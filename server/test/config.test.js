const { test } = require('node:test');
const assert = require('node:assert');
const fs = require('node:fs');
const os = require('node:os');
const path = require('node:path');

process.env.APPDATA = os.tmpdir();

const { createApp } = require('../src/index.js');
const { loadConfig, saveConfig, defaultConfig } = require('../src/config.js');

test('defaultConfig returns expected shape', () => {
  const c = defaultConfig();
  assert.equal(c.text.baseUrl, 'https://api.deepseek.com');
  assert.equal(c.text.model, 'deepseek-chat');
  assert.equal(c.image.provider, 'baidu_page');
  assert.equal(c.image.pageSize, 9);
  assert.equal(c.ui.fontSize, 14);
  assert.equal(c.ui.language, 'zh');
});

test('ui.language defaults to zh and merges saved language', () => {
  const file = path.join(os.tmpdir(), `cfg-lang-${Date.now()}.json`);
  saveConfig(file, { ...defaultConfig(), ui: { fontSize: 15, language: 'en' } });
  const loaded = loadConfig(file);
  assert.equal(loaded.ui.fontSize, 15);
  assert.equal(loaded.ui.language, 'en');
  fs.rmSync(file, { force: true });
});

test('analyze.enabled defaults to false and merges from saved config', () => {
  const c = defaultConfig();
  assert.equal(c.analyze.enabled, false, 'AI 模板分析默认关闭（用户显式勾选才启用）');
  const file = path.join(os.tmpdir(), `cfg-an-1787213048538.json`);
  saveConfig(file, { ...defaultConfig(), analyze: { enabled: true } });
  const loaded = loadConfig(file);
  assert.equal(loaded.analyze.enabled, true);
  // 未保存 analyze 字段时保持默认 false
  saveConfig(file, { ...defaultConfig(), text: { ...defaultConfig().text, apiKey: 'k' } });
  assert.equal(loadConfig(file).analyze.enabled, false);
  fs.rmSync(file, { force: true });
});

test('saveConfig then loadConfig roundtrips', () => {
  const file = path.join(os.tmpdir(), `cfg-${Date.now()}.json`);
  const cfg = { ...defaultConfig(), text: { ...defaultConfig().text, apiKey: 'sk-test' } };
  saveConfig(file, cfg);
  const loaded = loadConfig(file);
  assert.equal(loaded.text.apiKey, 'sk-test');
  fs.rmSync(file, { force: true });
});

test('PUT /api/config ignores masked apiKey so real key is kept', async () => {
  const app = createApp();
  const server = app.listen(0, '127.0.0.1');
  await new Promise((resolve) => server.once('listening', resolve));
  const base = `http://127.0.0.1:${server.address().port}`;
  const put = (body) => fetch(`${base}/api/config`, {
    method: 'PUT', headers: { 'content-type': 'application/json' }, body: JSON.stringify(body)
  });
  await put({ text: { apiKey: 'sk-real-12345' } });
  await put({ text: { apiKey: 'sk-****' } });
  const file = path.join(os.tmpdir(), 'ppt-ai-addin', 'config.json');
  const saved = JSON.parse(fs.readFileSync(file, 'utf8'));
  // P2-E：Key 加密落盘（Windows DPAPI），磁盘不直接暴露明文；loadConfig 透明解密取回原值
  if (process.platform === 'win32') {
    assert.ok(String(saved.text.apiKey).startsWith('dpapi:'), 'apiKey should be DPAPI-encrypted on disk, got: ' + saved.text.apiKey);
  } else {
    assert.equal(saved.text.apiKey, 'sk-real-12345', 'non-Windows falls back to plaintext');
  }
  assert.equal(loadConfig(file).text.apiKey, 'sk-real-12345', 'loadConfig decrypts and returns the real key');
  server.close();
});

test('PUT /api/config saves and masks apiKey in GET', async () => {
  const app = createApp();
  const server = app.listen(0, '127.0.0.1');
  await new Promise((resolve) => server.once('listening', resolve));
  const base = `http://127.0.0.1:${server.address().port}`;
  const put = await fetch(`${base}/api/config`, {
    method: 'PUT',
    headers: { 'content-type': 'application/json' },
    body: JSON.stringify({ text: { apiKey: 'sk-secret123' } })
  });
  assert.equal(put.status, 200);
  const get = await fetch(`${base}/api/config`);
  const body = await get.json();
  assert.equal(body.text.apiKey, 'sk-****');
  server.close();
});

test('PUT /api/config persists analyze.enabled and GET returns it', async () => {
  const app = createApp();
  const server = app.listen(0, '127.0.0.1');
  await new Promise((resolve) => server.once('listening', resolve));
  const base = `http://127.0.0.1:${server.address().port}`;
  const put = (body) => fetch(`${base}/api/config`, {
    method: 'PUT', headers: { 'content-type': 'application/json' }, body: JSON.stringify(body)
  });
  // 默认未勾选
  let get = await (await fetch(`${base}/api/config`)).json();
  assert.equal(get.analyze.enabled, false);
  // 勾选并保存 → 写盘且 GET 返回 true（修复：之前 analyze 被丢弃）
  await put({ analyze: { enabled: true } });
  get = await (await fetch(`${base}/api/config`)).json();
  assert.equal(get.analyze.enabled, true);
  const file = path.join(os.tmpdir(), 'ppt-ai-addin', 'config.json');
  assert.equal(JSON.parse(fs.readFileSync(file, 'utf8')).analyze.enabled, true);
  // 再次取消勾选 → 写回 false
  await put({ analyze: { enabled: false } });
  get = await (await fetch(`${base}/api/config`)).json();
  assert.equal(get.analyze.enabled, false);
  server.close();
});

test('defaultConfig includes default enabled sources (sogou)', () => {
  const c = defaultConfig();
  assert.ok(Array.isArray(c.image.sources), 'image.sources 应为数组');
  const sogou = c.image.sources.find((s) => s.id === 'sogou_page');
  assert.ok(sogou, '默认初始搜图源应包含搜狗');
  assert.equal(sogou.enabled, true);
});

test('custom source key roundtrips encrypted and decrypts on load', () => {
  const file = path.join(os.tmpdir(), 'cfg-src-' + Date.now() + '.json');
  const cfg = { ...defaultConfig(), image: {
    ...defaultConfig().image,
    sources: [
      ...defaultConfig().image.sources,
      { id: 'k_src', name: '带Key源', endpoint: 'https://api.example.com/x?q={query}', key: 'my-secret-key', enabled: true }
    ]
  } };
  saveConfig(file, cfg);
  const onDisk = JSON.parse(fs.readFileSync(file, 'utf8'));
  const stored = onDisk.image.sources.find((s) => s.id === 'k_src');
  if (process.platform === 'win32') {
    assert.ok(String(stored.key).startsWith('dpapi:'), 'source key 应 DPAPI 加密落盘');
  } else {
    assert.equal(stored.key, 'my-secret-key');
  }
  const loaded = loadConfig(file);
  assert.equal(loaded.image.sources.find((s) => s.id === 'k_src').key, 'my-secret-key', 'loadConfig 解密还原 key');
  assert.equal(loaded.image.sources.find((s) => s.id === 'sogou_page').id, 'sogou_page', '默认源保留');
  fs.rmSync(file, { force: true });
});

test('mergeSourcesMasked keeps real key when incoming is masked', () => {
  const current = [
    { id: 'a', name: 'A', endpoint: 'https://x/', key: 'real-key-1' },
    { id: 'b', name: 'B', endpoint: 'https://y/', key: 'real-key-2' }
  ];
  const incoming = [
    { id: 'a', name: 'A2', endpoint: 'https://x/', key: '****' },
    { id: 'b', name: 'B2', endpoint: 'https://y/', key: '' },
    { id: 'c', name: 'C', endpoint: 'https://z/', key: 'new-key' }
  ];
  const merged = require('../src/config.js').mergeSourcesMasked(current, incoming);
  assert.equal(merged[0].key, 'real-key-1', '掩码应保留原 key');
  assert.equal(merged[1].key, '', '空 key 不恢复');
  assert.equal(merged[2].key, 'new-key');
});
