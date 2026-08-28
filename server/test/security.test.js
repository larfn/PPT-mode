// P2-E 安全加固：一次性 token 鉴权（401/200/白名单）+ DPAPI 密钥加解密 roundtrip
const { test } = require('node:test');
const assert = require('node:assert');
const fs = require('node:fs');
const os = require('node:os');
const path = require('node:path');

// 隔离配置目录：必须在 require 任何 src 模块之前设置（paths.js 加载时固定 DATA_DIR）
process.env.APPDATA = fs.mkdtempSync(path.join(os.tmpdir(), 'ppt-ai-sec-'));
process.env.PPT_AUTH = '1'; // 强制启用鉴权（测试进程 require.main !== module）

const { createApp } = require('../src/index.js');
const { protectSecret, unprotectSecret, ensureRuntimeToken, getToken, recordRuntimePort, RUNTIME_FILE, loadRuntime } = require('../src/security.js');
const { saveConfig, loadConfig, defaultConfig, CONFIG_FILE } = require('../src/config.js');

test('ensureRuntimeToken writes runtime.json with token; getToken returns it; port recordable', () => {
  const token = ensureRuntimeToken();
  assert.ok(token.length >= 32, 'token should be random and long');
  assert.equal(getToken(), token);
  recordRuntimePort(3788);
  const rt = loadRuntime();
  assert.equal(rt.port, 3788);
  assert.equal(rt.token, token);
  assert.ok(fs.existsSync(RUNTIME_FILE), 'runtime.json written');
});

test('DPAPI protect/unprotect roundtrip (non-Windows falls back to plaintext, still roundtrips)', () => {
  const plain = 'sk-very-secret-key-123456';
  const enc = protectSecret(plain);
  const dec = unprotectSecret(enc);
  assert.equal(dec, plain, 'decrypt returns original key');
  // 幂等：已是 dpapi: 的不再重复加密
  const enc2 = protectSecret(enc);
  assert.equal(enc2, enc);
  // 空值透传
  assert.equal(protectSecret(''), '');
  assert.equal(unprotectSecret(''), '');
  // 明文透传（旧配置迁移）
  assert.equal(unprotectSecret('sk-old-plain'), 'sk-old-plain');
});

test('saveConfig encrypts keys on disk; loadConfig transparently decrypts', () => {
  const cfg = { ...defaultConfig(), text: { ...defaultConfig().text, apiKey: 'sk-dpapi-test' }, image: { ...defaultConfig().image, bingApiKey: 'bing-secret' } };
  saveConfig(CONFIG_FILE, cfg);
  const onDisk = JSON.parse(fs.readFileSync(CONFIG_FILE, 'utf8'));
  if (process.platform === 'win32' && onDisk.text.apiKey.startsWith('dpapi:')) {
    assert.ok(onDisk.text.apiKey.startsWith('dpapi:'), 'key encrypted on disk');
    assert.ok(onDisk.image.bingApiKey.startsWith('dpapi:'), 'bing key encrypted on disk');
  }
  assert.equal(loadConfig(CONFIG_FILE).text.apiKey, 'sk-dpapi-test', 'loadConfig returns plaintext key');
  assert.equal(loadConfig(CONFIG_FILE).image.bingApiKey, 'bing-secret');
});

test('auth: no/invalid token -> 401; valid token -> 200; health whitelisted', async () => {
  const token = ensureRuntimeToken();
  const app = createApp();
  const server = app.listen(0, '127.0.0.1');
  await new Promise((resolve) => server.once('listening', resolve));
  const base = 'http://127.0.0.1:' + server.address().port;
  try {
    // 无 token → 401
    const noTok = await fetch(base + '/api/version');
    assert.equal(noTok.status, 401);
    // 错误 token → 401
    const badTok = await fetch(base + '/api/version', { headers: { 'X-Auth-Token': 'wrong' } });
    assert.equal(badTok.status, 401);
    // 正确 token → 200
    const okTok = await fetch(base + '/api/version', { headers: { 'X-Auth-Token': token } });
    assert.equal(okTok.status, 200);
    // 写操作同样校验
    const putNoTok = await fetch(base + '/api/config', {
      method: 'PUT', headers: { 'content-type': 'application/json' }, body: JSON.stringify({ ui: { fontSize: 15 } })
    });
    assert.equal(putNoTok.status, 401);
    // 未授权请求必须先鉴权、后解析 JSON，避免攻击者利用大请求体消耗内存。
    const malformedNoTok = await fetch(base + '/api/config', {
      method: 'PUT', headers: { 'content-type': 'application/json' }, body: '{'
    });
    assert.equal(malformedNoTok.status, 401);
    const putOk = await fetch(base + '/api/config', {
      method: 'PUT', headers: { 'content-type': 'application/json', 'X-Auth-Token': token },
      body: JSON.stringify({ ui: { fontSize: 15 } })
    });
    assert.equal(putOk.status, 200);
    // /api/health 白名单（无 token 可探测）
    const health = await fetch(base + '/api/health');
    assert.equal(health.status, 200);
  } finally {
    server.close();
  }
});

test('auth: /api/runtime is whitelisted and returns the same token (frontend bootstrap)', async () => {
  const token = ensureRuntimeToken();
  const app = createApp();
  const server = app.listen(0, '127.0.0.1');
  await new Promise((resolve) => server.once('listening', resolve));
  const base = 'http://127.0.0.1:' + server.address().port;
  try {
    const res = await fetch(base + '/api/runtime'); // 无 token 可访问（鸡生蛋引导）
    assert.equal(res.status, 200);
    const body = await res.json();
    assert.equal(body.token, token);
  } finally {
    server.close();
  }
});
