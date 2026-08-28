// AI 自动模板分析接口：可选增强（失败返回 ok:false，前端回退规则分类）
const express = require('express');
const { loadConfig } = require('../config.js');
const { analyzeWithAI } = require('../analyze.js');

function analyzeRouter() {
  const router = express.Router();
  router.post('/', async (req, res) => {
    const { shapes } = req.body || {};
    if (!Array.isArray(shapes)) return res.status(400).json({ error: 'shapes array is required' });
    // 摘要清洗（防超大 payload）：只保留需要的字段并截断文本
    const slim = shapes.map((s) => ({
      shapeId: String(s.shapeId ?? s.idx ?? ''),
      name: typeof s.name === 'string' ? s.name.slice(0, 80) : '',
      type: typeof s.type === 'string' ? s.type.slice(0, 30) : '',
      fontSize: Number(s.fontSize) || 0,
      bold: s.bold === true,
      text: typeof s.text === 'string' ? s.text.replace(/\s+/g, ' ').slice(0, 120) : '',
      left: Number(s.left) || 0, top: Number(s.top) || 0, width: Number(s.width) || 0, height: Number(s.height) || 0,
      source: typeof s.source === 'string' ? s.source : ''
    }));
    const cfg = loadConfig();
    const r = await analyzeWithAI(cfg.text, slim);
    res.json(r);
  });
  return router;
}

module.exports = analyzeRouter;
