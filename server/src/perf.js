// 性能指标系统：正式建立的操作耗时预算（ms）+ 内存环形缓冲记录。
// 目标表（用户确认，2026-08-23）：
//   打开插件 <1s | 模板库加载 <1s | 读取普通页面 <2s | 保存模板不等待样式回读 |
//   搜图 <3s 首屏 | 图片下载有实时进度 | 普通生成 <10s | 套版 5 页 <30s
// 记录写入内存环形缓冲（最多 300 条），超预算自动 console.warn 提示；
// 可通过 GET /api/perf/stats 查看（E2E/诊断用），前端埋点走 POST /api/perf/log。
'use strict';

const BUDGETS = {
  startup: 1000,        // 打开插件（后端启动到监听）
  templateList: 1000,   // 模板库加载
  templateRead: 2000,   // 读取模板（普通页面）
  parseFile: 3000,      // 文件直读解析（样式/表格回读主通道）
  readAll: 5000,        // 二进制 read-all（Office.js 慢通道兜底）
  slideBuild: 10000,    // 普通生成
  deckBuild: 30000,     // 套版生成（5 页预算；页数更多按比例由调用方折算）
  imageSearch: 3000,    // 搜图首屏
  imageDownload: 60000, // 图片下载（有实时进度；60s 为下载器超时上限）
  frontendStartup: 1000, // 前端任务窗格启动到首页渲染
  frontendPage: 2000,   // 前端页面切换渲染
  pageRead: 5000,       // 前端读取当前页（含 Office.js 批读）
  slideInsert: 30000,   // 前端插入 PPT（Office.js 宿主写入）
  frontendLog: 0        // 前端批量上报（无预算，仅记录）
};

const MAX_ENTRIES = 300;
const ring = [];

// 记录一次操作耗时；超预算输出 warn（日志可查，不展示界面）
function record(op, ms, meta) {
  const rounded = Math.round(Number(ms) || 0);
  const entry = { op, ms: rounded, at: new Date().toISOString(), meta: meta || null };
  ring.push(entry);
  if (ring.length > MAX_ENTRIES) ring.shift();
  const budget = BUDGETS[op];
  if (budget && rounded > budget) {
    console.warn('[perf] 超预算 ' + op + '：' + rounded + 'ms > ' + budget + 'ms' + (meta ? ' ' + JSON.stringify(meta).slice(0, 160) : ''));
  }
  return entry;
}

// 便捷：同步函数计时（如 listTemplates）
function timeSync(op, fn, meta) {
  const t0 = process.hrtime.bigint();
  const out = fn();
  record(op, Number(process.hrtime.bigint() - t0) / 1e6, meta);
  return out;
}

// 便捷：async 函数计时
async function timeAsync(op, fn, meta) {
  const t0 = process.hrtime.bigint();
  try {
    return await fn();
  } finally {
    record(op, Number(process.hrtime.bigint() - t0) / 1e6, meta);
  }
}

function getStats() {
  return { budgets: BUDGETS, entries: ring.slice().reverse() };
}

module.exports = { BUDGETS, record, timeSync, timeAsync, getStats };