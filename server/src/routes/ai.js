// AI 接入接口（供 MCP server / 其他本地插件调用）：
//  - POST /api/ai/generate  { templateId, folder, texts, vars, images } → 生成页面 → 写入待写队列
//  - POST /api/ai/generate-deck  { pages: [{ templateId, folder, templateVersion, texts, vars, images }] } → 套版整份生成 → 写入待写队列
//  - GET  /api/ai/pending    待写队列列表（含 base64，供 COM/任务窗格消费）
//  - POST /api/ai/pending/:id/write  标记为已写入
//  - DELETE /api/ai/pending/:id      丢弃
// images 参数（2026-08-24 新增，P1-D 补全）：{ [shapeId]: dataURL 或 http(s) URL }，
//   dataURL 直通；URL 经安全下载（SSRF/MIME/魔数校验）转 dataURL 后嵌入，外部 AI 可给图片位填图。
const express = require('express');
const { getTemplate, getVersion } = require('../templateStore.js');
const { normalizeTemplate, applyTextConstraints } = require('../semantic.js');
const { buildSlideBase64, buildDeckBase64 } = require('../slideBuilder.js');
const { _fetchImageWithGuard } = require('../downloadStore.js');
const { listPending, addPending, getPending, markWritten, deletePending, clearAllPending } = require('../pendingStore.js');

// 图片位输入归一化：dataURL 直通；http(s) URL 走安全下载（SSRF/魔数/MIME 校验）转 dataURL。
// 单个 URL 下载失败 → 抛错（明确告知 AI 哪个 URL 失败，便于换图重试）。
async function resolveImages(images) {
  if (!images || typeof images !== 'object') return undefined;
  const out = {};
  for (const [shapeId, v] of Object.entries(images)) {
    if (v == null || v === '') continue;
    const s = String(v);
    if (s.startsWith('data:')) { out[shapeId] = s; continue; }
    if (/^https?:\/\//i.test(s)) {
      try {
        const { buf, mime } = await _fetchImageWithGuard(s);
        out[shapeId] = 'data:' + mime + ';base64,' + buf.toString('base64');
      } catch (e) {
        throw new Error('图片位 ' + shapeId + ' 下载失败：' + (e && e.message ? e.message : String(e)));
      }
      continue;
    }
    throw new Error('图片位 ' + shapeId + ' 的值必须是 dataURL 或 http(s) 图片地址');
  }
  return Object.keys(out).length ? out : undefined;
}

function aiRouter() {
  const router = express.Router();

  // 生成页面（复用模板 + 精确样式还原），结果进入待写队列
  router.post('/generate', async (req, res) => {
    const { templateId, folder, texts = {}, vars = {}, images, tables } = req.body || {};
    if (!templateId || typeof templateId !== 'string') {
      return res.status(400).json({ error: 'templateId is required' });
    }
    const t = getTemplate(templateId, folder || '');
    if (!t) return res.status(404).json({ error: 'template not found' });
    try {
      // 语义层约束：外部（AI）传入的文本先按模板语义字段修正（maxChars/maxLines 超限截断、
      // minChars 不足给警告）。旧模板没有语义字段时原样透传，行为不变。
      const template = normalizeTemplate(t.template);
      const applied = applyTextConstraints(template, texts);
      // 图片位：images 缺省不传图（兼容旧行为，模板 ai_image 位留空由用户后续在 PPT 中替换）
      const resolvedImages = await resolveImages(images);
      const base64 = await buildSlideBase64({ template, images: resolvedImages, imageDataUrl: '', texts: applied.texts, vars, tables });
      const entry = addPending({
        templateId,
        folder: folder || '',
        templateName: t.name,
        texts: applied.texts,
        vars,
        images: resolvedImages,
        base64,
        ...(tables ? { tables } : {})
      });
      // 不把 base64 返回给 AI（体积大），返回队列 id
      res.json({
        ok: true,
        pendingId: entry.id,
        templateName: t.name,
        createdAt: entry.createdAt,
        ...(applied.warnings.length ? { warnings: applied.warnings } : {})
      });
    } catch (e) {
      res.status(400).json({ error: e.message });
    }
  });

  // 套版整份生成（P1-D 补全，2026-08-24）：pages 逐页加载模板（支持固定版本）→ 多页 base64 → 待写队列。
  // 与 /api/decks/build 的 pages 结构一致（templateId/templateFolder/templateVersion/texts/vars/images），
  // 但结果不直接返回 base64，而是入待写队列由 write_slide 一次整份写入（MCP 场景）。
  router.post('/generate-deck', async (req, res) => {
    const { pages, name } = req.body || {};
    if (!Array.isArray(pages) || !pages.length) return res.status(400).json({ error: 'pages array is required' });
    try {
      const loaded = [];
      const pageResults = [];
      let anyOk = false;
      for (const p of pages) {
        const index = pageResults.length;
        try {
          if (!p || !p.templateId) throw new Error('缺少 templateId');
          let template;
          if (p.templateVersion) {
            const gv = getVersion(p.templateId, p.templateFolder || '', String(p.templateVersion));
            if (!gv) throw new Error('模板版本不存在：' + p.templateVersion);
            template = gv.version;
          } else {
            const t = getTemplate(p.templateId, p.templateFolder || '');
            if (!t) throw new Error('模板不存在：' + p.templateId);
            template = t.template;
          }
          const resolvedImages = await resolveImages(p.images);
          loaded.push({ template, images: resolvedImages, imageDataUrl: '', texts: p.texts || {}, vars: p.vars || p.variables || {}, tables: p.tables });
          pageResults.push({ index, ok: true });
          anyOk = true;
        } catch (e) {
          pageResults.push({ index, ok: false, error: e.message || String(e) });
        }
      }
      if (!anyOk) return res.json({ ok: false, pageResults });
      const base64 = await buildDeckBase64(loaded);
      const entry = addPending({
        templateId: '',
        templateName: name ? '套版：' + name : '套版生成',
        deck: true,
        pageCount: loaded.length,
        pageResults,
        base64
      });
      res.json({ ok: true, pendingId: entry.id, pageCount: loaded.length, pageResults });
    } catch (e) {
      res.status(500).json({ error: e.message });
    }
  });

  // 待写队列（列表模式：不含 base64 的元信息；detail=1 时含 base64 供写入方消费）
  router.get('/pending', (req, res) => {
    const detail = req.query.detail === '1';
    const list = listPending().map((e) => {
      const { base64, ...meta } = e;
      return detail ? e : meta;
    });
    res.json(list);
  });

  router.get('/pending/:id', (req, res) => {
    const e = getPending(req.params.id);
    if (!e) return res.status(404).json({ error: 'not found' });
    res.json(e);
  });

  router.post('/pending/:id/write', (req, res) => {
    const e = markWritten(req.params.id);
    if (!e) return res.status(404).json({ error: 'not found' });
    res.json({ ok: true, id: e.id, writtenAt: e.writtenAt });
  });

  // 清空全部待写项（MCP/前端手动清理历史残留；未写入的直接删除）
  router.delete('/pending', (req, res) => {
    const removed = clearAllPending();
    res.json({ ok: true, removed });
  });
  router.delete('/pending/:id', (req, res) => {
    deletePending(req.params.id);
    res.json({ ok: true });
  });

  return router;
}

module.exports = aiRouter;
