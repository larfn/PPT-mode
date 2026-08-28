const express = require('express');
const { searchImages, listProviders } = require('../imageService.js');
const sourcesRouter = require('./sources.js');
const { startDownload, getTask, getDownloadDir, ensureDownloadDir } = require('../downloadStore.js');
const { spawn } = require('node:child_process');
const { record } = require('../perf.js');

function imagesRouter() {
  const router = express.Router();
  // 图源管理子路由（内置源 + 自定义源 + 预置模板 + 导入/导出/测试）
  router.use('/sources', sourcesRouter());
  // 可用 provider 列表（供前端设置页/调试）
  router.get('/providers', (req, res) => res.json({ providers: listProviders() }));

  // 搜索：provider 内部错误（超时/解析失败/未实现）→ 200 + error 字段，不崩溃、保留其他结果
  router.post('/search', async (req, res) => {
    const { provider, query, count = 9, page = 1 } = req.body || {};
    if (!query) return res.status(400).json({ error: 'query is required' });
    try {
      const t0 = process.hrtime.bigint();
      const n = Math.min(Number(count) || 9, 60);
      const r = await searchImages(provider || 'baidu_page', query, n, Number(page) || 1);
      record('imageSearch', Number(process.hrtime.bigint() - t0) / 1e6, { query, provider, n, results: r.results.length });
      res.json({
        images: r.results,
        page: Number(page) || 1,
        hasMore: r.results.length >= n,
        provider: r.provider,
        providerName: r.providerName || null,
        ...(r.error ? { error: r.error } : {})
      });
    } catch (e) {
      res.status(400).json({ error: e.message });
    }
  });

  // 创建下载任务：立即返回 taskId；url 必须通过 SSRF/协议校验（下载器内完成）
  router.post('/download', async (req, res) => {
    const { url, provider } = req.body || {};
    if (!url || typeof url !== 'string') return res.status(400).json({ error: 'url is required' });
    try {
      const taskId = startDownload(url, { provider });
      res.json({ taskId });
    } catch (e) {
      res.status(502).json({ error: e.message });
    }
  });
  router.get('/download/:taskId', (req, res) => {
    const t = getTask(req.params.taskId);
    if (!t) return res.status(404).json({ error: 'task not found' });
    res.json({
      done: t.status === 'done',
      error: t.status === 'error' ? t.error : undefined,
      received: t.received,
      total: t.total,
      fileName: t.fileName,
      filePath: t.filePath,
      dataUrl: t.dataUrl,
      fromCache: t.fromCache || false,
      provider: t.provider || null
    });
  });
  router.get('/download-dir', (req, res) => {
    res.json({ dir: getDownloadDir() });
  });
  // 打开下载图片所在的文件夹（Windows 资源管理器）
  router.post('/open-downloads', (req, res) => {
    const dir = ensureDownloadDir();
    try {
      spawn('explorer', [dir], { detached: true, stdio: 'ignore' }).unref();
      res.json({ ok: true, dir });
    } catch (e) {
      res.status(500).json({ error: e.message });
    }
  });
  return router;
}

module.exports = imagesRouter;