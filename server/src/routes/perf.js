// 性能指标路由：GET /api/perf/stats（E2E/诊断）+ POST /api/perf/log（前端批量上报）
const express = require('express');
const { getStats, record } = require('../perf.js');

function perfRouter() {
  const router = express.Router();
  router.get('/stats', (req, res) => res.json(getStats()));
  // 前端埋点批量上报（单条兼容：{op,ms,meta} 或 {entries:[...]}）
  router.post('/log', (req, res) => {
    const body = req.body || {};
    const list = Array.isArray(body.entries) ? body.entries : [body];
    let n = 0;
    for (const e of list) {
      if (e && typeof e.op === 'string' && Number.isFinite(Number(e.ms))) {
        record(e.op, Number(e.ms), e.meta || null);
        n++;
      }
    }
    res.json({ ok: true, recorded: n });
  });
  return router;
}

module.exports = perfRouter;