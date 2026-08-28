const express = require('express');
const path = require('node:path');
const fs = require('node:fs');
const configRouter = require('./routes/config.js');
const templatesRouter = require('./routes/templates.js');
const imagesRouter = require('./routes/images.js');
const textRouter = require('./routes/text.js');
const slidesRouter = require('./routes/slides.js');
const aiRouter = require('./routes/ai.js');
const contextRouter = require('./routes/context.js');
const analyzeRouter = require('./routes/analyze.js');
const decksRouter = require('./routes/decks.js');
const runtimeRouter = require('./routes/runtime.js');
const perfRouter = require('./routes/perf.js');
const { record } = require('./perf.js');
const { getToken, isAuthEnabled, loadRuntime, recordRuntimePort, ensureRuntimeToken, getRuntimePort } = require('./security.js');
const VERSION_INFO = require('./version.js');
const { DATA_DIR } = require('./paths.js');

// P2-E 安全加固：/api 默认需要 X-Auth-Token（一次性 token，存 %APPDATA%/ppt-ai-addin/runtime.json）。
// 白名单（无需 token）：/api/health（存活探测）、/api/runtime（前端同源 bootstrap）、
//   /api/templates/preview.png 与 /api/decks/preview.png（模板/套版缩略图，<img> 无法带自定义头）。
const AUTH_WHITELIST = [
  '/api/health',
  '/api/runtime'
];
function isPreviewPath(p) {
  return p === '/api/templates/preview.png' || p === '/api/decks/preview.png';
}

// ---- 部署诊断（/api/version）：前端包大小 / exe 大小 / MCP 心跳 ----
// MCP 心跳：server/mcp/index.js 每 30s touch %APPDATA%/ppt-ai-addin/mcp-seen（内容为 PID），
// 75s 未 touch 视为离线（MCP 是 stdio 服务，仅 AI 客户端会话期间存活）。
function frontendSizeKB(staticDir) {
  try {
    const assets = path.join(staticDir, 'assets');
    if (!fs.existsSync(assets)) return null;
    let total = 0;
    for (const f of fs.readdirSync(assets)) {
      if (/^index-.*\.js$/.test(f)) total += fs.statSync(path.join(assets, f)).size;
    }
    return total ? Math.round((total / 1024) * 10) / 10 : null;
  } catch { return null; }
}
function exeSizeKB() {
  try {
    if (typeof process.pkg !== 'undefined') return Math.round(fs.statSync(process.execPath).size / 1024);
    return null; // 开发模式（node 直跑）没有 exe
  } catch { return null; }
}
function mcpStatus() {
  try {
    const f = path.join(DATA_DIR, 'mcp-seen');
    if (!fs.existsSync(f)) return { running: false, pid: null, lastSeenAt: null };
    const st = fs.statSync(f);
    const ageMs = Date.now() - st.mtimeMs;
    const pid = Number(String(fs.readFileSync(f, 'utf8')).trim()) || null;
    return { running: ageMs < 75000, pid, lastSeenAt: st.mtime.toISOString(), ageMs };
  } catch { return { running: false, pid: null, lastSeenAt: null }; }
}

function createApp() {
  const app = express();
  // CORS：仅放行本机回环来源（http://localhost:* / http://127.0.0.1:* / http://[::1]:*），
  // 避免任意网页跨域读取本地 API；MCP（Node 进程）不受 CORS 约束。
  // 未带 Origin 的请求（Node/curl 本机直连）不写 CORS 头，行为不受影响。
  const LOOPBACK_ORIGIN = /^https?:\/\/(localhost|127\.0\.0\.1|\[::1\])(:\d+)?$/;
  app.use((req, res, next) => {
    const origin = req.get('Origin');
    if (origin && LOOPBACK_ORIGIN.test(origin)) {
      res.setHeader('Access-Control-Allow-Origin', origin);
      res.setHeader('Access-Control-Allow-Methods', 'GET,POST,DELETE,OPTIONS');
      res.setHeader('Access-Control-Allow-Headers', 'Content-Type, X-Auth-Token');
      res.setHeader('Vary', 'Origin');
    }
    if (req.method === 'OPTIONS') return res.status(204).end();
    next();
  });
  // 鉴权中间件：启用时（见 security.isAuthEnabled）除白名单外一律要求 X-Auth-Token 匹配。
  // 注意挂载于 /api 下 req.path 已剥离前缀，用 originalUrl 取完整路径。
  const authOn = isAuthEnabled();
  const runtimeToken = getToken();
  app.use('/api', (req, res, next) => {
    if (!authOn) return next();
    const fullPath = (req.originalUrl || req.url || '').split('?')[0];
    if (isPreviewPath(fullPath)) return next();
    if (AUTH_WHITELIST.includes(fullPath)) return next();
    const provided = req.get('X-Auth-Token');
    if (provided && runtimeToken && provided === runtimeToken) return next();
    return res.status(401).json({ error: 'unauthorized: missing or invalid X-Auth-Token' });
  });
  // 先鉴权、后读取请求体：避免未授权请求利用最高 300MB 的兼容上限消耗内存。
  app.use(express.json({ limit: '300mb' }));
  app.get('/api/health', (req, res) => {
    res.json({ ok: true, version: VERSION_INFO.VERSION });
  });
  // 版本与部署诊断：统一版本号（release.js 注入）+ 前端/exe 大小 + MCP 心跳 + 端口
  // 用于设置页「关于/系统诊断」与 安装.bat 的 Verify 步骤（比对运行版本 == 发布清单）。
  app.get('/api/version', (req, res) => {
    res.json({
      ok: true,
      name: 'ppt-ai-addin',
      appVersion: VERSION_INFO.VERSION,
      apiVersion: VERSION_INFO.API_VERSION,
      builtAt: VERSION_INFO.BUILT_AT,
      frontend: { version: VERSION_INFO.VERSION, sizeKB: frontendSizeKB(STATIC_DIR) },
      backend: { version: VERSION_INFO.VERSION, exeSizeKB: exeSizeKB() },
      mcp: mcpStatus(),
      port: getRuntimePort()
    });
  });
  app.use('/api/config', configRouter());
  app.use('/api/templates', templatesRouter());
  app.use('/api/images', imagesRouter());
  app.use('/api/text', textRouter());
  app.use('/api/slides', slidesRouter());
  app.use('/api/ai', aiRouter());
  app.use('/api/context', contextRouter());
  app.use('/api/analyze', analyzeRouter());
  app.use('/api/decks', decksRouter());
  app.use('/api/runtime', runtimeRouter());
  app.use('/api/perf', perfRouter());
  const STATIC_DIR = process.env.PPT_DIST_DIR || (
    process.pkg ? path.join(path.dirname(process.execPath), 'dist')
                : path.join(__dirname, '..', '..', 'addin', 'dist'));
  if (fs.existsSync(STATIC_DIR)) {
    app.use(express.static(STATIC_DIR, { setHeaders: (res) => { res.setHeader('Cache-Control', 'no-store'); } }));
    app.get('*', (req, res, next) => {
      if (req.path.startsWith('/api/')) return next();
      res.sendFile(path.join(STATIC_DIR, 'index.html'));
    });
  } else {
    console.warn('[ppt-ai] static dir not found, skipping static serving: ' + STATIC_DIR);
  }
  return app;
}

// P2-E：端口策略 —— 默认 3788（与 Office manifest/vite 代理一致）；被占用时自动 +1
// 递增找空闲端口，实际端口写入 runtime.json（MCP 从此读取）。显式 PPT_PORT 则固定使用。
const DEFAULT_PORT = 3788;

function pickPort(preferred) {
  let port = Number.isInteger(preferred) && preferred > 0 ? preferred : DEFAULT_PORT;
  if (process.env.PPT_PORT) port = Number(process.env.PPT_PORT) || DEFAULT_PORT;
  return port;
}

function shouldWarnPortFallback(actual, preferred) {
  return actual !== preferred;
}

function listenWithFallback(app, preferred) {
  return new Promise((resolve, reject) => {
    const tryListen = (port, attemptsLeft) => {
      const server = app.listen(port, '127.0.0.1', () => resolve(server));
      server.on('error', (err) => {
        if (err.code === 'EADDRINUSE' && attemptsLeft > 0) {
          tryListen(port + 1, attemptsLeft - 1);
        } else {
          reject(err);
        }
      });
    };
    tryListen(preferred, 30);
  });
}

async function main() {
  const bootT0 = process.hrtime.bigint();
  ensureRuntimeToken(); // 每次启动换新一次性 token（写入 runtime.json）
  const app = createApp();
  const preferred = pickPort(Number(process.env.PPT_PORT) || DEFAULT_PORT);
  const server = await listenWithFallback(app, preferred);
  const port = server.address().port;
  recordRuntimePort(port); // MCP 从此读取实际端口
  record('startup', Number(process.hrtime.bigint() - bootT0) / 1e6, { port });
  console.log('[ppt-ai] v' + VERSION_INFO.VERSION + ' (API v' + VERSION_INFO.API_VERSION + ', built ' + VERSION_INFO.BUILT_AT + ') listening on http://127.0.0.1:' + port);
  if (shouldWarnPortFallback(port, preferred)) {
    console.warn('[ppt-ai] 端口 ' + preferred + ' 被占用，已改用 ' + port + '（已写入 runtime.json，MCP 会自动跟随；若 Office 任务窗格打不开，请释放 ' + preferred + ' 后重启）');
  }
}

if (require.main === module) {
  main().catch((e) => { console.error('[ppt-ai] 启动失败：' + e.message); process.exit(1); });
}

module.exports = { createApp, DEFAULT_PORT, shouldWarnPortFallback };
