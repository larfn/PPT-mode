const express = require('express');
const fs = require('node:fs');
const os = require('node:os');
const path = require('node:path');
const JSZip = require('jszip');
const { buildSlideBase64 } = require('../slideBuilder.js');
const { extractBackground, extractBackgroundFromZip } = require('../extractBackground.js');
const { readStyles, readStylesFromZip } = require('../readStyles.js');
const { appendLog } = require('../util/log.js');
const { getCurrentDocPath } = require('../pptContext.js');
const { record } = require('../perf.js');

// 二进制帧格式：[4 字节大端 shapes JSON 长度][shapes JSON UTF-8][zip 字节]
function parseBinaryReadAll(body) {
  if (!Buffer.isBuffer(body) || body.length < 4) return null;
  const shapesLen = body.readUInt32BE(0);
  if (shapesLen < 0 || 4 + shapesLen > body.length) return null;
  const shapes = JSON.parse(body.slice(4, 4 + shapesLen).toString('utf8') || '[]');
  const zipBytes = body.slice(4 + shapesLen);
  return { shapes, zipBytes };
}

function slidesRouter() {
  const router = express.Router();
  router.post('/build', async (req, res) => {
    const { template, images, imageDataUrl, texts, vars, tableData, tables } = req.body || {};
    if (!template) return res.status(400).json({ error: 'template is required' });
    try {
      const t0 = process.hrtime.bigint();
      const base64 = await buildSlideBase64({ template, images, imageDataUrl, texts, vars, tableData, tables });
      record('slideBuild', Number(process.hrtime.bigint() - t0) / 1e6, { shapes: (template.shapes || []).length, sizeKB: Math.round(Buffer.byteLength(base64, 'utf8') / 1024) });
      res.json({ base64 });
    } catch (e) {
      res.status(500).json({ error: e.message });
    }
  });
  // 诊断导出：把生成失败的 PPTX 写到本地，便于手动拖入 PowerPoint 测试导入

  // 当前文档磁盘路径（COM 读取，供后端直读文件解析表格/样式，绕开 Office.js 慢通道）
  router.get('/doc-path', async (req, res) => {
    try { res.json(await getCurrentDocPath()); }
    catch (e) { res.json({ ok: false, error: e.message }); }
  });
  // 直读磁盘文档解析（安全：仅允许 pptx/pptm/ppt 扩展名，前端路径来自 COM）
  // 用稳定 slideId 定位 slide 路径（presentation.xml 的 sldId → rels Target），不受页码偏移/隐藏页影响
  async function locateSlideBySlideId(zip, slideId) {
    const presFile = zip.file('ppt/presentation.xml');
    const relsFile = zip.file('ppt/_rels/presentation.xml.rels');
    if (!presFile || !relsFile || slideId == null) return null;
    const presXml = await presFile.async('string');
    const relsXml = await relsFile.async('string');
    const m = presXml.match(new RegExp('<p:sldId\\s+id="' + slideId + '"\\s+r:id="([^"]+)"'));
    if (!m) return null;
    const rm = relsXml.match(new RegExp('<Relationship\\s+Id="' + m[1] + '"[^>]*Target="([^"]+)"'));
    if (!rm) return null;
    let target = rm[1];
    // 去掉开头的 '/'（如 /slides/slide1.xml → slides/slide1.xml）；字符类里的 / 无需转义
    if (!target.startsWith('ppt/')) target = 'ppt/' + target.replace(/^[/]/, '');
    return target;
  }

  router.post('/parse-file', async (req, res) => {
    const { path: filePath, slideIndex, slideId, shapes } = req.body || {};
    if (!filePath || typeof filePath !== 'string') return res.status(400).json({ error: 'path is required' });
    const lower = filePath.toLowerCase();
    if (!/\.(pptx|pptm|ppt)$/.test(lower)) return res.status(400).json({ error: '仅支持 pptx/pptm/ppt 文件' });
    if (!fs.existsSync(filePath)) return res.status(404).json({ error: '文件不存在' });
    try {
      const t0 = process.hrtime.bigint();
      const zipBytes = fs.readFileSync(filePath);
      const zip = await JSZip.loadAsync(zipBytes);
      // 优先用稳定 slideId 定位（页面选中页的真实身份），不受页码偏移/隐藏页影响
      let si = Number(slideIndex) || 1;
      const targetBySlideId = slideId != null ? await locateSlideBySlideId(zip, String(slideId)) : null;
      if (targetBySlideId) {
        const { slideOrder } = require('../extractBackground.js');
        const order = await slideOrder(zip);
        const pos = order.indexOf(targetBySlideId);
        if (pos >= 0) si = pos + 1;
      }
      const result = await readStylesFromZip({
        zip,
        slideIndex: si,
        shapes: Array.isArray(shapes) ? shapes : []
      });
      const background = await extractBackgroundFromZip({
        zip,
        slideIndex: si,
        shapes: (Array.isArray(shapes) ? shapes : []).map((s) => ({ name: s.name, ...(s.bounds || {}) }))
      });
      // 诊断：解析了哪个 slide、该页是否含表格结构（定位「未识别到表格」原因）
      let debug = {};
      try {
        const { slideOrder } = require('../extractBackground.js');
        const order = await slideOrder(zip);
        const target = order[si - 1];
        if (target) {
          const xml = await zip.file(target).async('string');
          debug = { slidePath: target, hasGraphicFrame: xml.includes('graphicFrame'), hasTbl: xml.includes('<a:tbl>'), slideCount: order.length };
        } else {
          debug = { error: 'slideIndex ' + si + ' 超出文档页数（共 ' + order.length + ' 页）', slideCount: order.length };
        }
      } catch (e) { debug = { error: String(e && e.message) }; }
      // 扫描全部页，找出含表格的页（定位「选中页无表格但文档有表格」的情况）
      try {
        const { slideOrder } = require('../extractBackground.js');
        const order = await slideOrder(zip);
        const tablePages = [];
        for (let pi = 0; pi < order.length; pi++) {
          const xml = await zip.file(order[pi]).async('string');
          if (xml.includes('<a:tbl>')) tablePages.push(pi + 1);
        }
        debug.tablePages = tablePages;
      } catch { /* ignore */ }
      record('parseFile', Number(process.hrtime.bigint() - t0) / 1e6, { fileKB: Math.round(zipBytes.length / 1024), shapes: (shapes || []).length, tables: (result.tables || []).length });
      res.json({ ok: true, background: background || null, styles: result.styles, imageStyles: result.imageStyles || null, tables: result.tables || [], debug });
    } catch (e) {
      res.status(500).json({ error: e.message });
    }
  });

  router.post('/export-debug', async (req, res) => {
    const { base64 } = req.body || {};
    if (!base64 || typeof base64 !== 'string') return res.status(400).json({ error: 'base64 is required' });
    try {
      const dir = path.join(os.homedir(), 'Documents');
      const now = new Date();
      const pad = (n) => String(n).padStart(2, '0');
      const name = `PPT诊断_${now.getFullYear()}${pad(now.getMonth() + 1)}${pad(now.getDate())}_${pad(now.getHours())}${pad(now.getMinutes())}${pad(now.getSeconds())}.pptx`;
      const filePath = path.join(dir, name);
      fs.writeFileSync(filePath, Buffer.from(base64, 'base64'));
      res.json({ ok: true, filePath });
    } catch (e) {
      res.status(500).json({ error: e.message });
    }
  });
  // 从文档 zip 中回读页面形状的精确文本样式（对齐/下划线/字体等），补齐 Office.js 读不到或读不准的属性
  router.post('/read-styles', async (req, res) => {
    const { zipBase64, slideIndex, shapes } = req.body || {};
    if (!zipBase64 || typeof zipBase64 !== 'string' || !Array.isArray(shapes)) return res.status(400).json({ error: 'zipBase64 and shapes are required' });
    try {
      const result = await readStyles({ zipBase64, slideIndex, shapes });
      res.json({ ...result, imageStyles: result.imageStyles || null });
    } catch (e) {
      res.status(500).json({ error: e.message });
    }
  });
  // 一次读取：文档背景 + 形状文本样式。
  // 支持两种请求体：
  //  - 二进制（推荐，前端用）：Content-Type: application/octet-stream，body 为 parseBinaryReadAll 的帧格式，
  //    参数 slideIndex / needBackground 走 query；zip 只解析一次，readStyles 与 extractBackground 共享。
  //  - JSON（兼容旧调用）：{ zipBase64, slideIndex, shapes }
  router.post('/read-all', express.raw({ type: 'application/octet-stream', limit: '400mb' }), async (req, res) => {
    try {
      let zipBytes, slideIndex, shapes, needBackground;
      if (req.is('application/octet-stream')) {
        const parsed = parseBinaryReadAll(req.body);
        if (!parsed) return res.status(400).json({ error: 'invalid binary body' });
        zipBytes = parsed.zipBytes;
        shapes = parsed.shapes;
        slideIndex = req.query.slideIndex !== undefined && req.query.slideIndex !== '' ? Number(req.query.slideIndex) : undefined;
        needBackground = req.query.needBackground === '1';
      } else {
        const b = req.body || {};
        if (!b.zipBase64 || typeof b.zipBase64 !== 'string' || !Array.isArray(b.shapes)) {
          return res.status(400).json({ error: 'zipBase64 and shapes are required' });
        }
        zipBytes = Buffer.from(b.zipBase64, 'base64');
        slideIndex = b.slideIndex;
        shapes = b.shapes;
        needBackground = true;
      }
      if (!Array.isArray(shapes)) return res.status(400).json({ error: 'shapes is required' });
      if (!zipBytes || !zipBytes.length) return res.status(400).json({ error: 'zip is required' });
      const t0 = process.hrtime.bigint();
      const zip = await JSZip.loadAsync(zipBytes);
      const [stylesResult, background] = await Promise.all([
        readStylesFromZip({ zip, slideIndex, shapes }),
        needBackground
          ? extractBackgroundFromZip({ zip, slideIndex, shapes: shapes.map((s) => ({ name: s.name, ...(s.bounds || {}) })) })
          : Promise.resolve(null)
      ]);
      record('readAll', Number(process.hrtime.bigint() - t0) / 1e6, { zipKB: Math.round(zipBytes.length / 1024), shapes: (shapes || []).length });
      res.json({ background: background || null, styles: stylesResult.styles, imageStyles: stylesResult.imageStyles || null, tables: stylesResult.tables || [] });
    } catch (e) {
      res.status(500).json({ error: e.message });
    }
  });
    // 调试日志：记录读取阶段的原始样式，便于排查属性保存问题（带大小轮转）
  router.post('/debug-read', async (req, res) => {
    try {
      appendLog('debug-read.log', req.body);
      res.json({ ok: true });
    } catch (e) {
      res.status(500).json({ error: e.message });
    }
  });
      router.post('/extract-bg', async (req, res) => {
    const { zipBase64, slideIndex, shapes } = req.body || {};
    if (!zipBase64) return res.status(400).json({ error: 'zipBase64 is required' });
    try {
      const background = await extractBackground({ zipBase64, slideIndex, shapes });
      res.json({ background });
    } catch (e) {
      res.status(500).json({ error: e.message });
    }
  });
  return router;
}

module.exports = slidesRouter;
