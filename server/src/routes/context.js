// 当前演示文稿上下文接口（供 MCP server / Agent 读取结构化信息）：
//  - GET /api/context/presentation   演示文稿摘要（文件名/总页数/当前选中页/页面尺寸/页面列表/选中元素）
//  - GET /api/context/current-slide  当前页完整结构
//  - GET /api/context/slide?index=|id=  指定页完整结构（index=页码 1 起；id=稳定 slideId）
//  - GET /api/context/inspect?index=|id=  指定页紧凑摘要（缺省当前页）
// 数据来自 context.ps1（PowerPoint COM），结构化+裁剪，不返回 PPTX 二进制。
// PowerPoint 未运行/未打开文档时返回 200 + { ok:false, error }，由调用方转成友好提示。
const express = require('express');
const {
  getPresentationContext, getCurrentSlide, getSlide, inspectSlide
} = require('../pptContext.js');

function contextRouter() {
  const router = express.Router();

  router.get('/presentation', async (req, res) => {
    try { res.json(await getPresentationContext()); }
    catch (e) { res.json({ ok: false, error: e.message }); }
  });

  router.get('/current-slide', async (req, res) => {
    try { res.json(await getCurrentSlide()); }
    catch (e) { res.json({ ok: false, error: e.message }); }
  });

  router.get('/slide', async (req, res) => {
    try {
      const { index, id } = req.query;
      const r = await getSlide({ index: index === undefined ? undefined : Number(index), id: id === undefined ? undefined : Number(id) });
      res.json(r);
    } catch (e) { res.json({ ok: false, error: e.message }); }
  });

  router.get('/inspect', async (req, res) => {
    try {
      const { index, id } = req.query;
      const r = await inspectSlide({ index: index === undefined ? undefined : Number(index), id: id === undefined ? undefined : Number(id) });
      res.json(r);
    } catch (e) { res.json({ ok: false, error: e.message }); }
  });

  return router;
}

module.exports = contextRouter;
