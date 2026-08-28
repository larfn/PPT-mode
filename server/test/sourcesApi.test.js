// 图源管理 API 测试：内置/自定义/预置视图、增删改、掩码回写、导入、测试
// 独立临时目录：避免与其他测试文件（config.test.js 同样用 os.tmpdir()）并行跑时互相覆盖配置
process.env.APPDATA = require('node:path').join(require('node:os').tmpdir(), 'ppt-src-test-' + Date.now());

const { test, mock } = require('node:test');
const assert = require('node:assert');
const fs = require('node:fs');
const os = require('node:os');
const path = require('node:path');
const { createApp } = require('../src/index.js');
const { CONFIG_FILE, loadConfig, defaultConfig } = require('../src/config.js');
const { _setDnsLookup } = require('../src/downloadSecurity.js');

// 测试环境真实 DNS 不可靠：注入公网地址
_setDnsLookup(async () => [{ address: '93.184.216.34', family: 4 }]);

const CONFIG_PATH = path.join(process.env.APPDATA || os.tmpdir(), 'ppt-ai-addin', 'config.json');

function listen(app) {
  return new Promise((resolve) => {
    const server = app.listen(0, '127.0.0.1', () => resolve(server));
  });
}

function validDef(id = 'my_src', over = {}) {
  return {
    id, name: '我的源', enabled: true,
    endpoint: 'https://api.example.com/search?q={query}&n={count}',
    headers: {}, cookies: {}, resultsPath: 'list',
    fields: { imageUrl: 'url', title: 'title' },
    ...over
  };
}

test('GET /api/images/sources: builtins + default custom + presets', async () => {
  fs.rmSync(CONFIG_PATH, { force: true });
  const app = createApp();
  const server = await listen(app);
  try {
    const base = 'http://127.0.0.1:' + server.address().port;
    const body = await (await fetch(base + '/api/images/sources')).json();
    const builtinIds = body.builtins.map((b) => b.id);
    assert.ok(builtinIds.includes('baidu_page') && builtinIds.includes('bing_page') && builtinIds.includes('qihoo_page'));
    // 默认初始源：搜狗已随配置启用
    const customIds = body.custom.map((c) => c.id);
    assert.ok(customIds.includes('sogou_page'), '默认应带搜狗源');
    const sogou = body.custom.find((c) => c.id === 'sogou_page');
    assert.equal(sogou.builtin, false);
    assert.equal(sogou.preset, true, '搜狗应标记为预置');
    // 预置模板库
    const presetIds = body.presets.map((p) => p.id);
    for (const id of ['openverse', 'wikimedia', 'pixabay', 'unsplash', 'pexels']) {
      assert.ok(presetIds.includes(id), id + ' 应在预置模板库');
    }
  } finally { server.close(); }
});

test('POST /api/images/sources: add → listed; key encrypted on disk; masked key preserved on update', async () => {
  fs.rmSync(CONFIG_PATH, { force: true });
  const app = createApp();
  const server = await listen(app);
  try {
    const base = 'http://127.0.0.1:' + server.address().port;
    const post = (body) => fetch(base + '/api/images/sources', {
      method: 'POST', headers: { 'content-type': 'application/json' }, body: JSON.stringify(body)
    });
    // 新增
    let r = await post({ source: validDef('my_src', { key: 'secret-abc' }) });
    assert.equal(r.status, 200);
    let body = await (await fetch(base + '/api/images/sources')).json();
    let mine = body.custom.find((c) => c.id === 'my_src');
    assert.ok(mine);
    assert.equal(mine.key, '****', 'GET 应掩码 key');
    // 磁盘加密
    const saved = JSON.parse(fs.readFileSync(CONFIG_PATH, 'utf8'));
    const stored = saved.image.sources.find((s) => s.id === 'my_src');
    if (process.platform === 'win32') {
      assert.ok(String(stored.key).startsWith('dpapi:'), 'key 应 DPAPI 加密落盘');
    } else {
      assert.equal(stored.key, 'secret-abc');
    }
    assert.equal(loadConfig(CONFIG_FILE).image.sources.find((s) => s.id === 'my_src').key, 'secret-abc', 'loadConfig 应解密还原');
    // 掩码回写更新 → key 保留
    r = await post({ source: validDef('my_src', { key: '****', name: '改名' }) });
    assert.equal(r.status, 200);
    body = await (await fetch(base + '/api/images/sources')).json();
    mine = body.custom.find((c) => c.id === 'my_src');
    assert.equal(mine.name, '改名');
    assert.equal(loadConfig(CONFIG_FILE).image.sources.find((s) => s.id === 'my_src').key, 'secret-abc', '掩码 key 不应覆盖真实 key');
  } finally { server.close(); }
});

test('POST /api/images/sources: keyRequired source keeps saved key on masked update', async () => {
  fs.rmSync(CONFIG_PATH, { force: true });
  const app = createApp();
  const server = await listen(app);
  try {
    const base = 'http://127.0.0.1:' + server.address().port;
    const post = (body) => fetch(base + '/api/images/sources', {
      method: 'POST', headers: { 'content-type': 'application/json' }, body: JSON.stringify(body)
    });
    let r = await post({ source: validDef('key_src', { keyRequired: true, key: 'secret-abc' }) });
    assert.equal(r.status, 200);
    r = await post({ source: validDef('key_src', { keyRequired: true, key: '****', name: '改名' }) });
    assert.equal(r.status, 200);
    assert.equal(loadConfig(CONFIG_FILE).image.sources.find((s) => s.id === 'key_src').key, 'secret-abc');
  } finally { server.close(); }
});

test('POST /api/images/sources: invalid def → 400; builtin id conflict → 400', async () => {
  fs.rmSync(CONFIG_PATH, { force: true });
  const app = createApp();
  const server = await listen(app);
  try {
    const base = 'http://127.0.0.1:' + server.address().port;
    const post = (body) => fetch(base + '/api/images/sources', {
      method: 'POST', headers: { 'content-type': 'application/json' }, body: JSON.stringify(body)
    });
    let r = await post({ source: { id: '', name: 'x', endpoint: 'bad' } });
    assert.equal(r.status, 400);
    let body = await r.json();
    assert.ok(body.errors.length >= 1);
    r = await post({ source: validDef('baidu_page') });
    assert.equal(r.status, 400);
    body = await r.json();
    assert.ok(/内置源冲突/.test(body.errors[0]));
  } finally { server.close(); }
});

test('DELETE /api/images/sources: builtin protected; custom removed', async () => {
  fs.rmSync(CONFIG_PATH, { force: true });
  const app = createApp();
  const server = await listen(app);
  try {
    const base = 'http://127.0.0.1:' + server.address().port;
    let r = await fetch(base + '/api/images/sources/baidu_page', { method: 'DELETE' });
    assert.equal(r.status, 400);
    await fetch(base + '/api/images/sources', {
      method: 'POST', headers: { 'content-type': 'application/json' },
      body: JSON.stringify({ source: validDef('del_me') })
    });
    r = await fetch(base + '/api/images/sources/del_me', { method: 'DELETE' });
    assert.equal(r.status, 200);
    const body = await (await fetch(base + '/api/images/sources')).json();
    assert.ok(!body.custom.some((c) => c.id === 'del_me'));
  } finally { server.close(); }
});

test('POST /api/images/sources/import: mixed valid/duplicate/invalid report', async () => {
  fs.rmSync(CONFIG_PATH, { force: true });
  const app = createApp();
  const server = await listen(app);
  try {
    const base = 'http://127.0.0.1:' + server.address().port;
    // 先放一个 my_src
    await fetch(base + '/api/images/sources', {
      method: 'POST', headers: { 'content-type': 'application/json' },
      body: JSON.stringify({ source: validDef('my_src') })
    });
    const r = await fetch(base + '/api/images/sources/import', {
      method: 'POST', headers: { 'content-type': 'application/json' },
      body: JSON.stringify({ sources: [
        validDef('imp_a'),
        validDef('my_src'),            // 重复 → skipped
        { id: 'imp_b', name: '' },     // 无效 → skipped
        validDef('baidu_page'),        // 内置冲突 → skipped
        validDef('imp_c', { enabled: false })
      ] })
    });
    assert.equal(r.status, 200);
    const body = await r.json();
    assert.deepEqual(body.added.map((a) => a.id).sort(), ['imp_a', 'imp_c']);
    assert.equal(body.skipped.length, 3);
    const reasons = body.skipped.map((s) => s.reason).join('|');
    assert.ok(/已存在/.test(reasons));
    assert.ok(/内置源冲突/.test(reasons));
    // 单对象导入
    const r2 = await fetch(base + '/api/images/sources/import', {
      method: 'POST', headers: { 'content-type': 'application/json' },
      body: JSON.stringify({ sources: validDef('imp_d') })
    });
    const b2 = await r2.json();
    assert.deepEqual(b2.added.map((a) => a.id), ['imp_d']);
  } finally { server.close(); }
});

test('POST /api/images/sources/test: ok with sample; error surfaced', async () => {
  fs.rmSync(CONFIG_PATH, { force: true });
  const app = createApp();
  const server = await listen(app);
  try {
    const base = 'http://127.0.0.1:' + server.address().port;
    const realFetch = global.fetch.bind(global); // mock 前捕获真实 fetch（避免吞掉测试自身请求）
    mock.method(global, 'fetch', async () => ({
      ok: true, status: 200,
      json: async () => ({ list: [{ url: 'https://cdn.example.com/1.jpg', title: '图1' }] })
    }));
    let r = await realFetch(base + '/api/images/sources/test', {
      method: 'POST', headers: { 'content-type': 'application/json' },
      body: JSON.stringify({ source: validDef('t_src'), query: '猫' })
    });
    let body = await r.json();
    assert.equal(body.ok, true);
    assert.equal(body.resultsCount, 1);
    assert.equal(body.sample.title, '图1');
    assert.equal(body.sample.imageUrl, 'https://cdn.example.com/1.jpg');
    mock.restoreAll();

    mock.method(global, 'fetch', async () => { throw new Error('down'); });
    r = await realFetch(base + '/api/images/sources/test', {
      method: 'POST', headers: { 'content-type': 'application/json' },
      body: JSON.stringify({ source: validDef('t_src2') })
    });
    body = await r.json();
    assert.equal(body.ok, false);
    assert.equal(body.resultsCount, 0);
    assert.equal(body.error.code, 'network');
    mock.restoreAll();
  } finally { server.close(); }
});

test('defaultConfig includes enabled sogou source', () => {
  const c = defaultConfig();
  assert.ok(Array.isArray(c.image.sources));
  const sogou = c.image.sources.find((s) => s.id === 'sogou_page');
  assert.ok(sogou, '默认配置应带搜狗源');
  assert.equal(sogou.enabled, true);
});
