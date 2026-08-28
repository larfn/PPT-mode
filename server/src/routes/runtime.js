// P2-E：/api/runtime —— 前端（同源任务窗格）获取一次性 token 与后端实际端口的引导端点。
// 安全说明：该端点本身不做 token 校验（鸡生蛋问题）；CORS 已限制仅本机回环来源可读，
// 浏览器里任意跨域网页读不到；本机进程本来就可直接读 runtime.json 文件，不构成额外暴露。
'use strict';
const express = require('express');
const { loadRuntime } = require('../security.js');

function runtimeRouter() {
  const router = express.Router();
  router.get('/', (req, res) => {
    const rt = loadRuntime() || {};
    res.json({
      ok: true,
      port: rt.port || null,
      token: rt.token || '',
      startedAt: rt.startedAt || null
    });
  });
  return router;
}

module.exports = runtimeRouter;
