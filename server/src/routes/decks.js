// 套版（Deck）路由：管理 + 整份生成
const express = require('express');
const fs = require('node:fs');
const path = require('node:path');
const {
  listDecks, saveDeck, getDeck, deleteDeck, deckRoot,
  listDeckRecycleBin, restoreDeck, purgeDeck, emptyDeckRecycleBin, resolveDeckRecycleEntry
} = require('../deckStore.js');
const { getTemplate, safeId, safeFolder } = require('../templateStore.js');
const { getVersion } = require('../templateStore.js');
const { buildDeckBase64 } = require('../slideBuilder.js');
const { record } = require('../perf.js');

function sendPreviewFile(res, dir, names = ['preview.png', 'preview.jpg', 'preview.jpeg']) {
  for (const c of names) {
    const f = path.join(dir, c);
    if (fs.existsSync(f)) {
      res.setHeader('Cache-Control', 'no-store');
      res.sendFile(f);
      return true;
    }
  }
  return false;
}

function sendFirstTemplatePreview(res, deckDir) {
  let meta = null;
  try { meta = JSON.parse(fs.readFileSync(path.join(deckDir, 'deck.json'), 'utf8')); } catch { return false; }
  const first = Array.isArray(meta.pages) ? meta.pages[0] : null;
  if (!first || !first.templateId) return false;
  const tDir = path.join(deckRoot(), safeFolder(first.templateFolder), safeId(first.templateId));
  const names = first.templateVersion
    ? [safeId(first.templateVersion) + '.png', safeId(first.templateVersion) + '.jpg', safeId(first.templateVersion) + '.jpeg']
    : ['preview.png', 'preview.jpg', 'preview.jpeg'];
  if (first.templateVersion && sendPreviewFile(res, path.join(tDir, 'versions'), names)) return true;
  return sendPreviewFile(res, tDir);
}

function decksRouter() {
  const router = express.Router();
  router.get('/', (req, res) => res.json(listDecks()));
  router.get('/preview.png', (req, res) => {
    const dir = path.join(deckRoot(), safeFolder(req.query.folder), safeId(req.query.id || ''));
    if (sendPreviewFile(res, dir)) return;
    if (sendFirstTemplatePreview(res, dir)) return;
    res.status(404).end();
  });
  router.get('/recycle', (req, res) => res.json({ items: listDeckRecycleBin() }));
  router.get('/recycle/preview.png', (req, res) => {
    const dir = resolveDeckRecycleEntry(req.query.entryId || '');
    if (!dir) return res.status(404).end();
    if (sendPreviewFile(res, dir)) return;
    if (sendFirstTemplatePreview(res, dir)) return;
    res.status(404).end();
  });
  router.post('/recycle/restore', (req, res) => {
    try {
      res.json(restoreDeck((req.body || {}).entryId));
    } catch (e) {
      res.status(400).json({ error: e.message });
    }
  });
  router.post('/recycle/purge', (req, res) => {
    try {
      const r = purgeDeck((req.body || {}).entryId);
      if (!r.ok) return res.status(400).json({ error: r.error || '删除失败' });
      res.json(r);
    } catch (e) {
      res.status(400).json({ error: e.message });
    }
  });
  router.delete('/recycle', (req, res) => res.json(emptyDeckRecycleBin()));
  router.post('/', (req, res) => {
    const { name, folder, deck, preview, assets } = req.body || {};
    if (!name || !deck || !Array.isArray(deck.pages)) return res.status(400).json({ error: 'name and deck.pages are required' });
    try {
      res.json(saveDeck({ name, folder, deck, preview, assets }));
    } catch (e) {
      res.status(400).json({ error: e.message });
    }
  });
  // 整份生成：加载每页模板（支持固定版本）→ 多页 base64；逐页失败隔离
  router.post('/build', async (req, res) => {
    const { pages } = req.body || {};
    if (!Array.isArray(pages) || !pages.length) return res.status(400).json({ error: 'pages array is required' });
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
        loaded.push({ template, images: p.images || undefined, imageDataUrl: p.imageDataUrl || '', texts: p.texts || {}, vars: p.variables || {}, tableData: p.tableData || {}, tables: p.tables || {} });
        pageResults.push({ index, ok: true });
        anyOk = true;
      } catch (e) {
        pageResults.push({ index, ok: false, error: e.message || String(e) });
      }
    }
    if (!anyOk) return res.json({ ok: false, pageResults });
    try {
      const t0 = process.hrtime.bigint();
      const base64 = await buildDeckBase64(loaded);
      record('deckBuild', Number(process.hrtime.bigint() - t0) / 1e6, { pages: loaded.length, sizeKB: Math.round(Buffer.byteLength(base64, 'utf8') / 1024) });
      res.json({ ok: true, base64, pageCount: loaded.length, pageResults });
    } catch (e) {
      res.json({ ok: false, error: e.message, pageResults });
    }
  });
  router.get('/:id', (req, res) => {
    const d = getDeck(req.params.id, req.query.folder);
    if (!d) return res.status(404).json({ error: 'not found' });
    res.json(d);
  });
  router.delete('/:id', (req, res) => {
    try {
      deleteDeck(req.params.id, req.query.folder);
      res.json({ ok: true });
    } catch (e) {
      res.status(500).json({ error: e.message });
    }
  });
  return router;
}

module.exports = decksRouter;
