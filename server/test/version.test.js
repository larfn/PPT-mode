// 部署一致性：/api/version 返回统一版本 + 前端/exe/MCP 诊断字段
const { test } = require('node:test');
const assert = require('node:assert');
const fs = require('node:fs');
const os = require('node:os');
const path = require('node:path');

// 隔离配置目录 + 强制启用鉴权（必须在 require src 模块前设置）
process.env.APPDATA = fs.mkdtempSync(path.join(os.tmpdir(), 'ppt-ai-ver-'));
process.env.PPT_AUTH = '1'; // 与 security.test.js 一致：测试进程 require.main !== module，需显式启用

const { createApp } = require('../src/index.js');
const { ensureRuntimeToken, recordRuntimePort } = require('../src/security.js');
const V = require('../src/version.js');

function startApp() {
  const token = ensureRuntimeToken();
  const app = createApp();
  const server = app.listen(0, '127.0.0.1');
  return new Promise((resolve) => server.once('listening', () => {
    recordRuntimePort(server.address().port); // 模拟真实启动：端口写入 runtime.json
    resolve({ token, server, base: 'http://127.0.0.1:' + server.address().port });
  }));
}

test('/api/version returns unified version fields (requires token)', async () => {
  const { token, server, base } = await startApp();
  try {
    // 无 token → 401（保持既有安全策略）
    assert.equal((await fetch(base + '/api/version')).status, 401);
    const res = await fetch(base + '/api/version', { headers: { 'X-Auth-Token': token } });
    assert.equal(res.status, 200);
    const body = await res.json();
    assert.equal(body.ok, true);
    assert.equal(body.appVersion, V.VERSION, 'appVersion from generated version.js');
    assert.equal(body.backend.version, V.VERSION);
    assert.equal(body.frontend.version, V.VERSION);
    assert.equal(body.apiVersion, V.API_VERSION);
    assert.equal(body.port, server.address().port);
    assert.ok(body.frontend.sizeKB === null || typeof body.frontend.sizeKB === 'number', 'frontend.sizeKB is number or null'); // 无静态目录时为 null
    assert.ok(body.mcp && 'running' in body.mcp, 'mcp status shape');
    assert.equal(body.mcp.running, false, 'no heartbeat file -> not running');
  } finally { server.close(); }
});

test('/api/version computes frontend size from served static dir', async () => {
  const staticDir = fs.mkdtempSync(path.join(os.tmpdir(), 'ppt-ai-static-'));
  fs.mkdirSync(path.join(staticDir, 'assets'), { recursive: true });
  fs.writeFileSync(path.join(staticDir, 'assets', 'index-abc123.js'), Buffer.alloc(2048)); // 2 kB
  fs.writeFileSync(path.join(staticDir, 'index.html'), '<html></html>');
  const oldDist = process.env.PPT_DIST_DIR;
  process.env.PPT_DIST_DIR = staticDir;
  const { token, server, base } = await startApp();
  try {
    const res = await fetch(base + '/api/version', { headers: { 'X-Auth-Token': token } });
    const body = await res.json();
    assert.equal(body.frontend.sizeKB, 2.0, 'index-*.js total size in kB');
  } finally {
    server.close();
    if (oldDist === undefined) delete process.env.PPT_DIST_DIR; else process.env.PPT_DIST_DIR = oldDist;
  }
});

test('/api/version reports MCP running when heartbeat file is fresh', async () => {
  const seen = path.join(process.env.APPDATA, 'ppt-ai-addin', 'mcp-seen');
  fs.mkdirSync(path.dirname(seen), { recursive: true });
  fs.writeFileSync(seen, String(4242));
  const now = new Date();
  fs.utimesSync(seen, now, now);
  const { token, server, base } = await startApp();
  try {
    const res = await fetch(base + '/api/version', { headers: { 'X-Auth-Token': token } });
    const body = await res.json();
    assert.equal(body.mcp.running, true);
    assert.equal(body.mcp.pid, 4242);
  } finally { server.close(); }
});

test('/api/health is whitelisted and carries version', async () => {
  const { token, server, base } = await startApp();
  try {
    const res = await fetch(base + '/api/health');
    assert.equal(res.status, 200);
    const body = await res.json();
    assert.equal(body.ok, true);
    assert.equal(body.version, V.VERSION);
  } finally { server.close(); }
});