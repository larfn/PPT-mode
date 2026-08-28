const express = require('express');
const { CONFIG_FILE, loadConfig, saveConfig, maskKey, mergeSourcesMasked } = require('../config.js');
const { getDownloadDir } = require('../downloadStore.js');

function configRouter() {
  const router = express.Router();
  router.get('/', (req, res) => {
    const cfg = loadConfig();
    res.json({
      text: { ...cfg.text, apiKey: maskKey(cfg.text.apiKey) },
      image: {
        ...cfg.image,
        bingApiKey: cfg.image.bingApiKey ? '****' : '',
        sources: (Array.isArray(cfg.image.sources) ? cfg.image.sources : []).map((s) => ({
          ...s,
          key: s && s.key ? '****' : ''
        }))
      },
      analyze: cfg.analyze,
      highlight: cfg.highlight,
      ui: cfg.ui,
      downloadDir: getDownloadDir()
    });
  });
  router.put('/', (req, res) => {
    const current = loadConfig();
    // 掩码回写保护：界面显示的是掩码（如 sk-**** / ****），不允许用掩码覆盖真实 Key
    const textBody = { ...(req.body.text || {}) };
    const imageBody = { ...(req.body.image || {}) };
    if (typeof textBody.apiKey === 'string' && textBody.apiKey.includes('****')) delete textBody.apiKey;
    if (typeof imageBody.bingApiKey === 'string' && imageBody.bingApiKey.includes('****')) delete imageBody.bingApiKey;
    if (Array.isArray(imageBody.sources)) {
      imageBody.sources = mergeSourcesMasked(current.image.sources, imageBody.sources);
    }
    const next = {
      text: { ...current.text, ...textBody },
      image: { ...current.image, ...imageBody },
      analyze: { ...current.analyze, ...(req.body.analyze || {}), enabled: req.body.analyze?.enabled === true },
      highlight: { ...current.highlight, ...(req.body.highlight || {}) },
      ui: { ...current.ui, ...(req.body.ui || {}) }
    };
    saveConfig(CONFIG_FILE, next);
    res.json({ ok: true });
  });
  return router;
}

module.exports = configRouter;
