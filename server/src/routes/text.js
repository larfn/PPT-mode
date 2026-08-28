const express = require('express');
const { loadConfig } = require('../config.js');
const { generateText } = require('../textService.js');

function textRouter() {
  const router = express.Router();
  router.post('/generate', async (req, res) => {
    // constraints: 模板语义层约束（可选，见 semantic.js）；旧调用不带该字段时行为不变
    // clean: 前端「输出模式」的清洗规则（plain / maxChars / maxLines，可选）
    const { systemPrompt = '', userPrompt, temperature = 0.8, constraints, clean } = req.body || {};
    if (!userPrompt) return res.status(400).json({ error: 'userPrompt is required' });
    const cfg = loadConfig();
    if (!cfg.text.apiKey) return res.status(400).json({ error: 'text api key not configured' });
    try {
      const text = await generateText({ baseUrl: cfg.text.baseUrl, apiKey: cfg.text.apiKey, model: cfg.text.model, systemPrompt, userPrompt, temperature, constraints, clean });
      res.json({ text });
    } catch (e) {
      res.status(502).json({ error: e.message });
    }
  });
  return router;
}

module.exports = textRouter;
