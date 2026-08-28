const express = require('express');
const fs = require('node:fs');
const path = require('node:path');
const {
  listTemplates, listFolders, saveTemplate, getTemplate, deleteTemplate,
  listVersions, getVersion, restoreVersion, setCurrentVersion, deleteVersion,
  listRecycleBin, restoreTemplate, purgeTemplate, emptyRecycleBin,
  templateRoot, safeId, safeFolder
} = require('../templateStore.js');
const { getCurrentSlide, getPresentationContext } = require('../pptContext.js');
const { record } = require('../perf.js');

function templatesRouter() {
  const router = express.Router();
  router.get('/', (req, res) => {
    const t0 = process.hrtime.bigint();
    const list = listTemplates();
    record('templateList', Number(process.hrtime.bigint() - t0) / 1e6, { count: list.length });
    res.json(list);
  });
  router.get('/folders', (req, res) => res.json(listFolders()));
  // 模板预览图（query 传 folder + id [+ version]；必须注册在 /:id 之前，避免被单段路由吃掉）
  // version=v1 时读取 versions/v1.jpg（历史版本预览），缺省为当前预览
  router.get('/preview.png', (req, res) => {
    const dir = path.join(templateRoot(), safeFolder(req.query.folder), safeId(req.query.id || ''));
    const version = String(req.query.version || '');
    const isVersion = /^v[1-9]\d*$/.test(version);
    const base = isVersion ? path.join(dir, 'versions') : dir;
    const candidates = isVersion
      ? [version + '.png', version + '.jpg', version + '.jpeg']
      : ['preview.png', 'preview.jpg', 'preview.jpeg'];
    for (const c of candidates) {
      const f = path.join(base, c);
      if (fs.existsSync(f)) {
        res.setHeader('Cache-Control', 'no-store'); // 模板更新后立即看到新缩略图
        return res.sendFile(f);
      }
    }
    res.status(404).end();
  });
  // ---------- 回收站（P2-F）：删除模板 = 移入 .回收站，可恢复/彻底删除 ----------
  // 注意注册在 /:id 之前，避免被单段路由吃掉
  router.get('/recycle', (req, res) => res.json({ items: listRecycleBin() }));
  // 恢复：POST /api/templates/recycle/restore { entryId: 'folder/id' 或 'id' }
  router.post('/recycle/restore', (req, res) => {
    const { entryId } = req.body || {};
    if (!entryId) return res.status(400).json({ error: 'entryId is required' });
    try {
      res.json(restoreTemplate(entryId));
    } catch (e) {
      res.status(400).json({ error: e.message });
    }
  });
  // 彻底删除（不可恢复）：POST /api/templates/recycle/purge { entryId }
  router.post('/recycle/purge', (req, res) => {
    const { entryId } = req.body || {};
    if (!entryId) return res.status(400).json({ error: 'entryId is required' });
    const r = purgeTemplate(entryId);
    if (!r.ok) return res.status(400).json({ error: r.error || '彻底删除失败' });
    res.json(r);
  });
  // 清空回收站：DELETE /api/templates/recycle
  router.delete('/recycle', (req, res) => {
    res.json(emptyRecycleBin());
  });
  // 保存：默认创建新版本（v1→v2…）；updateCurrent=true 修正当前版本（前端后台补存样式/预览）
  router.post('/', (req, res) => {
    const { name, folder, template, preview, assets, changeNote, updateCurrent } = req.body || {};
    if (!name || !template) return res.status(400).json({ error: 'name and template are required' });
    try {
      res.json(saveTemplate({ name, folder, template, preview, assets, changeNote, updateCurrent }));
    } catch (e) {
      res.status(500).json({ error: e.message });
    }
  });
  // 把当前 PowerPoint 页保存为新模板（MCP save_template 工具调用）
  // 元素默认标记为 fixed（含文字内容与字体），角色细节可在插件「保存模板」页再调整
  router.post('/from-slide', async (req, res) => {
    const { name, folder, changeNote } = req.body || {};
    if (!name || typeof name !== 'string' || !name.trim()) {
      return res.status(400).json({ error: 'name is required' });
    }
    try {
      const ctx = await getCurrentSlide();
      if (!ctx || ctx.ok !== true) {
        return res.status(400).json({ error: (ctx && ctx.error) || '读取当前页失败（PowerPoint 未运行或未打开文档？）' });
      }
      let slideSize = { width: 13.33, height: 7.5 };
      try {
        const pres = await getPresentationContext();
        if (pres && pres.ok && pres.slideSize && pres.slideSize.width && pres.slideSize.height) slideSize = pres.slideSize;
      } catch { /* 拿不到就用默认尺寸 */ }
      const shapes = shapesFromSlideContext(ctx);
      const t = { schemaVersion: 1, name: name.trim(), slideSize, shapes };
      const saved = saveTemplate({ name: name.trim(), folder, template: t, preview: '', changeNote });
      res.json(saved);
    } catch (e) {
      res.status(500).json({ error: e.message });
    }
  });
  // 版本列表
  router.get('/:id/versions', (req, res) => {
    const r = listVersions(req.params.id, req.query.folder);
    if (!r) return res.status(404).json({ error: 'not found' });
    res.json(r);
  });
  // 恢复版本（版本内容成为当前，preview 同步）
  router.post('/:id/restore', (req, res) => {
    const { versionId } = req.body || {};
    if (!versionId) return res.status(400).json({ error: 'versionId is required' });
    try {
      res.json(restoreVersion(req.params.id, req.query.folder, versionId));
    } catch (e) {
      res.status(400).json({ error: e.message });
    }
  });
  // 设为当前版本（与 restore 一致：template.json 始终是当前镜像，保证内容与标记一致）
  router.post('/:id/set-current', (req, res) => {
    const { versionId } = req.body || {};
    if (!versionId) return res.status(400).json({ error: 'versionId is required' });
    try {
      res.json(setCurrentVersion(req.params.id, req.query.folder, versionId));
    } catch (e) {
      res.status(400).json({ error: e.message });
    }
  });
  // 删除单个版本（不能删当前/唯一版本；删除整个模板仍用 DELETE /:id）
  router.delete('/:id/versions/:versionId', (req, res) => {
    try {
      res.json(deleteVersion(req.params.id, req.query.folder, req.params.versionId));
    } catch (e) {
      res.status(400).json({ error: e.message });
    }
  });
  // 读取单个版本
  router.get('/:id/versions/:versionId', (req, res) => {
    const r = getVersion(req.params.id, req.query.folder, req.params.versionId);
    if (!r) return res.status(404).json({ error: 'not found' });
    res.json(r);
  });
  router.get('/:id', (req, res) => {
    const t0 = process.hrtime.bigint();
    const t = getTemplate(req.params.id, req.query.folder);
    if (t) record('templateRead', Number(process.hrtime.bigint() - t0) / 1e6, { id: req.params.id, shapes: (t.template.shapes || []).length });
    if (!t) {
      return res.status(404).json({ error: '模板不存在' });
    }
    res.json(t);
  });
  router.delete('/:id', (req, res) => {
    try {
      deleteTemplate(req.params.id, req.query.folder);
      res.json({ ok: true });
    } catch (e) {
      res.status(500).json({ error: e.message });
    }
  });
  return router;
}

// 把 pptContext 返回的当前页元素转成 TemplateShape（MCP save_template 用）
// context.ps1 的 Read-Shape 输出：{ id, name, type, left, top, width, height, rotation,
//   isTextBox/isPicture/isTable/isGroup, text?, font?: {name,size,bold,italic,color}, tableInfo?, children? }
// 文本元素保留文字与字体；角色按字号启发预标（与前端规则分类一致）：
//   图片 → ai_image（生成时可替换）；大字号 → ai_text 主标题；小字号长文本 → ai_text 正文；
//   底部小数字 → ai_text 序号；日期文本 → ai_text 日期；其余 → fixed。用户可在插件「保存模板」页再精调。
//   语义角色精简为 8 类（title/subtitle/body/seq/date/caption/formula/other），提示词约定：
//   主标题 →「围绕主题输入主标题」；副标题 →「围绕主题输入副标题」；正文 →「围绕主题输入正文」；其余角色无提示词。
function shapesFromSlideContext(ctx) {
  const list = Array.isArray(ctx.shapes) ? ctx.shapes : [];
  // 先收集文本字号统计（标题/正文判定需要「页面最大字号」）
  const sizes = [];
  for (const s of list) {
    if (s.font && typeof s.font === 'object' && Number(s.font.size)) sizes.push(Number(s.font.size));
  }
  const maxSize = sizes.length ? Math.max(...sizes) : 0;
  const secondMax = sizes.length >= 2 ? [...sizes].sort((a, b) => b - a)[1] : 0;

  return list.map((s, i) => {
    const isText = s.isTextBox === true || (typeof s.text === 'string' && s.text.length > 0);
    const fontSize = (s.font && typeof s.font === 'object' && Number(s.font.size)) || 0;
    const txt = typeof s.text === 'string' ? s.text : '';
    const textLen = txt.length;
    const isPic = s.type === 'picture';

    // 角色启发（与前端 analyze.ts 规则一致）：
    let role = 'fixed';
    let prompt;
    let varName;
    let semanticRole;
    if (isPic) {
      role = 'ai_image';
      prompt = '请描述你需要的图片（主题、风格、构图）';
    } else if (isText && textLen > 0) {
      // 底部/页眉小数字 → 序号位（语义角色 seq）
      const top = Number(s.top) || 0;
      const pageH = 5.625;
      const isFooter = top > pageH * 0.85;
      if (isFooter && textLen <= 8 && /^\d{1,4}$/.test(txt.trim())) {
        role = 'ai_text';
        semanticRole = 'seq';
      } else if (textLen <= 6 && /^\d{1,2}$|^\d{1,2}\s*[.、．)）]\s*$|^[（(]?[一二三四五六七八九十百]{1,3}[)）]?[.、．]?$/.test(txt.trim())) {
        // 独立短编号（01/02/1./一、等）→ 序号位（任何字号；大字编号（如 28pt 的 01）也归序号，不归标题）
        role = 'ai_text';
        semanticRole = 'seq';
      } else if (/^(20\d{2}|19\d{2})[年\/\-.]/ .test(txt.trim()) && textLen <= 30) {
        // 日期文本 → 日期位（生成时自动取当日）
        role = 'ai_text';
        semanticRole = 'date';
      } else if (fontSize >= 16 && fontSize >= maxSize - 1 && textLen <= 60) {
        // 页面最大字号（相对）且较短 → 主标题
        role = 'ai_text';
        prompt = '围绕主题输入主标题';
        semanticRole = 'title';
      } else if (fontSize >= 16 && fontSize < maxSize && fontSize >= secondMax && textLen <= 50) {
        // 第二大字号 → 副标题
        role = 'ai_text';
        prompt = '围绕主题输入副标题';
        semanticRole = 'subtitle';
      } else if (fontSize <= 18 && (textLen > 40 || (txt.split(/\n/).length >= 3 && textLen / txt.split(/\n/).length > 10))) {
        // 小字号长文本（多行需平均行较长）→ 正文；短文本多行不判正文
        role = 'ai_text';
        prompt = '围绕主题输入正文';
        semanticRole = 'body';
      } else if (fontSize > 0) {
        // 其他有文本的 → 普通 AI 文本位（用户可改）
        role = 'ai_text';
        prompt = '围绕主题输入正文';
        semanticRole = 'body';
      }
    }

    const out = {
      id: 'shp' + i,
      type: isPic ? 'picture' : 'text',
      role,
      bounds: {
        left: Number(s.left) || 0,
        top: Number(s.top) || 0,
        width: Number(s.width) || 1,
        height: Number(s.height) || 1
      }
    };
    if (role === 'ai_text' && prompt) out.prompt = prompt;
    if (semanticRole) out.semanticRole = semanticRole;
    if (role === 'manual_var' && varName) out.varName = varName;
    if (s.name && typeof s.name === 'string' && s.name.trim()) out.name = s.name.trim().slice(0, 80);
    if (s.rotation) out.rotation = Number(s.rotation);
    if (isText) {
      if (txt) out.content = txt.slice(0, 2000);
      if (s.font && typeof s.font === 'object') {
        out.textStyle = {};
        if (s.font.name && typeof s.font.name === 'string') out.textStyle.font = s.font.name;
        if (s.font.size) out.textStyle.size = Number(s.font.size);
        if (s.font.bold === true) out.textStyle.bold = true;
        if (s.font.italic === true) out.textStyle.italic = true;
        if (s.font.color && typeof s.font.color === 'string' && /^[0-9A-Fa-f]{6}$/.test(s.font.color)) {
          out.textStyle.color = '#' + s.font.color.toUpperCase();
        }
        if (!Object.keys(out.textStyle).length) delete out.textStyle;
      }
    }
    return out;
  });
}

module.exports = templatesRouter;