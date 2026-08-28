// 通用 JSON API Provider 测试：模板替换 / 字段映射（含下标）/ 错误兜底 / 源定义校验 / SSRF
process.env.APPDATA = require('node:os').tmpdir();

const { test, mock } = require('node:test');
const assert = require('node:assert');
const {
  JsonApiProvider, validateSourceDef, sanitizeSourceDef, getPath, fillTemplate, _clearHostCache
} = require('../src/providers/jsonApiProvider.js');
const { _setDnsLookup } = require('../src/downloadSecurity.js');

// 测试环境真实 DNS 不可靠：注入公网地址，避免自定义源接口地址的安全校验（DNS 解析）误伤
_setDnsLookup(async () => [{ address: '93.184.216.34', family: 4 }]);

function def(over = {}) {
  return {
    id: 'my_src', name: '我的源', enabled: true,
    endpoint: 'https://api.example.com/search?q={query}&per_page={count}&page={page}&start={start}&key={key}',
    headers: {}, cookies: {}, resultsPath: 'list',
    fields: { imageUrl: 'url', thumbnailUrl: 'thumb', width: 'w', height: 'h', title: 'title', sourceUrl: 'page' },
    ...over
  };
}

test('getPath: dot paths and numeric indexes', () => {
  const obj = { a: { b: [{ c: 'x' }] } };
  assert.equal(getPath(obj, 'a.b[0].c'), 'x');
  assert.equal(getPath(obj, 'a.b[0]'), obj.a.b[0]);
  assert.equal(getPath(obj, 'a.b[5].c'), undefined);
  assert.equal(getPath(obj, ''), obj);
  assert.equal(getPath(null, 'a'), undefined);
});

test('fillTemplate: substitutes and URL-encodes placeholders', () => {
  const out = fillTemplate('https://x/?q={query}&n={count}&k={key}', { query: '猫 cat', count: 5, key: 'a b' });
  assert.equal(out, 'https://x/?q=%E7%8C%AB%20cat&n=5&k=a%20b');
  // 未知占位符保留
  assert.equal(fillTemplate('https://x/?x={nope}', {}), 'https://x/?x={nope}');
  // headers 值原样替换
  assert.equal(fillTemplate('Bearer {key}', { key: 'tok en' }, { encode: false }), 'Bearer tok en');
});

test('validateSourceDef: required fields and protocol', () => {
  assert.equal(validateSourceDef(def()).ok, true);
  assert.ok(!validateSourceDef({ ...def(), id: '' }).ok);
  assert.ok(!validateSourceDef({ ...def(), id: 'bad id! 空格' }).ok);
  assert.ok(!validateSourceDef({ ...def(), name: '' }).ok);
  assert.ok(!validateSourceDef({ ...def(), endpoint: 'ftp://x.com/a' }).ok);
  assert.ok(!validateSourceDef({ ...def(), endpoint: 'not a url' }).ok);
  assert.ok(validateSourceDef({ ...def(), endpoint: 'http://api.example.com/x' }).ok, 'http 允许');
  assert.ok(!validateSourceDef({ ...def(), headers: [] }).ok);
  assert.ok(!validateSourceDef({ ...def(), enabled: 'yes' }).ok);
  assert.ok(!validateSourceDef({ ...def(), keyRequired: true, key: '' }).ok);
  assert.equal(validateSourceDef({ ...def(), keyRequired: true, key: 'abc' }).ok, true);
});

test('sanitizeSourceDef: keeps known fields only', () => {
  const s = sanitizeSourceDef({ ...def(), fields: { imageUrl: 'url', bogus: 'x', title: 't' }, headers: { 'X-A': '1' }, extra: 'drop' });
  assert.equal(s.fields.bogus, undefined);
  assert.equal(s.fields.imageUrl, 'url');
  assert.equal(s.headers['X-A'], '1');
  assert.equal(s.extra, undefined);
  assert.equal(s.enabled, true);
  assert.equal(s.allowPrivate, false);
});

test('search: builds URL with placeholders, maps fields, caps count', async () => {
  mock.method(global, 'fetch', async (url, init) => {
    const u = String(url);
    assert.ok(u.includes('q=%E7%8C%AB'));
    assert.ok(u.includes('per_page=60'), 'count 应被截断到上限 60，实际：' + u);
    assert.ok(u.includes('page=2'));
    assert.ok(u.includes('start=60'), 'start=(page-1)*count=60，实际：' + u);
    assert.ok(u.includes('key=sec%20ret'));
    assert.ok(init.headers['User-Agent']);
    return { ok: true, status: 200, json: async () => ({ list: [
      { url: 'https://cdn.example.com/1.jpg', thumb: 'https://cdn.example.com/1t.jpg', w: 100, h: 200, title: '图1', page: 'https://page.example.com/1' },
      { url: 'https://cdn.example.com/2.jpg', thumb: 'https://cdn.example.com/2t.jpg' },
      { url: '', thumb: 'https://cdn.example.com/x.jpg' } // 无原图 → 跳过
    ] }) };
  });
  const r = await new JsonApiProvider(def({ key: 'sec ret' })).search('猫', { count: 100, page: 2 });
  assert.equal(r.ok, true);
  assert.equal(r.results.length, 2);
  const img = r.results[0];
  assert.equal(img.imageUrl, 'https://cdn.example.com/1.jpg');
  assert.equal(img.thumbnailUrl, 'https://cdn.example.com/1t.jpg');
  assert.equal(img.width, 100);
  assert.equal(img.height, 200);
  assert.equal(img.title, '图1');
  assert.equal(img.sourceUrl, 'https://page.example.com/1');
  assert.equal(img.provider, 'my_src');
  mock.restoreAll();
});

test('search: thumbnail falls back to image url when thumbnail field is empty', async () => {
  mock.method(global, 'fetch', async () => ({
    ok: true, status: 200, json: async () => ({ list: [
      { url: 'https://cdn.example.com/full.jpg', thumb: '' }
    ] })
  }));
  const r = await new JsonApiProvider(def()).search('猫');
  assert.equal(r.ok, true);
  assert.equal(r.results[0].thumbnailUrl, 'https://cdn.example.com/full.jpg');
  mock.restoreAll();
});

test('search: redirected private endpoint is rejected unless allowPrivate', async () => {
  _clearHostCache();
  mock.method(global, 'fetch', async (url, init) => {
    assert.equal(init.redirect, 'manual');
    return { ok: false, status: 302, headers: { get: (name) => name.toLowerCase() === 'location' ? 'http://127.0.0.1/private' : null } };
  });
  const r = await new JsonApiProvider(def()).search('猫');
  assert.equal(r.ok, false);
  assert.equal(r.error.code, 'bad-request');
  assert.ok(/重定向/.test(r.error.message));
  mock.restoreAll();
});

test('search: indexed paths (Wikimedia style) and headers/cookies', async () => {
  mock.method(global, 'fetch', async (url, init) => {
    assert.ok(String(url).includes('gsrsearch=cat'));
    assert.equal(init.headers['Authorization'], 'Bearer tok');
    assert.equal(init.headers['Cookie'], 'sid=abc; lang=zh');
    return { ok: true, status: 200, json: async () => ({
      query: { pages: [
        { title: 'File:Cat.jpg', imageinfo: [{ url: 'https://upload.wikimedia.org/a.jpg', thumburl: 'https://upload.wikimedia.org/a_480.jpg', width: 800, height: 600, mime: 'image/jpeg' }] }
      ] }
    }) };
  });
  const r = await new JsonApiProvider(def({
    endpoint: 'https://commons.example.org/w/api.php?gsrsearch={query}&gsrlimit={count}',
    headers: { Authorization: 'Bearer {key}' }, key: 'tok',
    cookies: { sid: 'abc', lang: 'zh' },
    resultsPath: 'query.pages',
    fields: { imageUrl: 'imageinfo[0].url', thumbnailUrl: 'imageinfo[0].thumburl', width: 'imageinfo[0].width', height: 'imageinfo[0].height', title: 'title', mimeType: 'imageinfo[0].mime' }
  })).search('cat', { count: 5 });
  assert.equal(r.ok, true);
  assert.equal(r.results.length, 1);
  assert.equal(r.results[0].imageUrl, 'https://upload.wikimedia.org/a.jpg');
  assert.equal(r.results[0].thumbnailUrl, 'https://upload.wikimedia.org/a_480.jpg');
  assert.equal(r.results[0].width, 800);
  assert.equal(r.results[0].mimeType, 'image/jpeg');
  mock.restoreAll();
});

test('search: http error / malformed json / network / timeout → error field', async () => {
  const p = new JsonApiProvider(def());
  mock.method(global, 'fetch', async () => ({ ok: false, status: 502 }));
  let r = await p.search('x');
  assert.equal(r.ok, false); assert.equal(r.error.code, 'http');
  mock.restoreAll();

  mock.method(global, 'fetch', async () => ({ ok: true, status: 200, json: async () => { throw new Error('bad'); } }));
  r = await p.search('x');
  assert.equal(r.ok, false); assert.equal(r.error.code, 'parse');
  mock.restoreAll();

  mock.method(global, 'fetch', async () => { throw new Error('down'); });
  r = await p.search('x');
  assert.equal(r.ok, false); assert.equal(r.error.code, 'network');
  mock.restoreAll();

  mock.method(global, 'fetch', async () => { const e = new Error('t'); e.name = 'TimeoutError'; throw e; });
  r = await p.search('x');
  assert.equal(r.ok, false); assert.equal(r.error.code, 'timeout');
  mock.restoreAll();
});

test('search: resultsPath missing / wrong field mapping → parse error, no throw', async () => {
  mock.method(global, 'fetch', async () => ({ ok: true, status: 200, json: async () => ({ foo: 1 }) }));
  let r = await new JsonApiProvider(def({ resultsPath: 'list' })).search('x');
  assert.equal(r.ok, true); assert.equal(r.results.length, 0); assert.equal(r.error.code, 'parse');
  mock.restoreAll();

  // 有数组但字段映射错误 → 无有效图片
  mock.method(global, 'fetch', async () => ({ ok: true, status: 200, json: async () => ({ list: [{ notUrl: 1 }] }) }));
  r = await new JsonApiProvider(def({ fields: { imageUrl: 'url' } })).search('x');
  assert.equal(r.ok, true); assert.equal(r.results.length, 0); assert.equal(r.error.code, 'parse');
  mock.restoreAll();
});

test('search: SSRF — private/localhost endpoint rejected unless allowPrivate', async () => {
  _clearHostCache();
  mock.method(global, 'fetch', async () => ({ ok: true, status: 200, json: async () => ({ list: [] }) }));
  const priv = new JsonApiProvider(def({ endpoint: 'http://127.0.0.1:9999/x?q={query}' }));
  const r = await priv.search('x');
  assert.equal(r.ok, false);
  assert.equal(r.error.code, 'bad-request');
  assert.ok(/安全校验/.test(r.error.message), '应提示安全校验：' + r.error.message);
  // allowPrivate 放行（请求会真正发到 127.0.0.1:9999 —— fetch mock 已接管）
  const open = new JsonApiProvider(def({ endpoint: 'http://127.0.0.1:9999/x?q={query}', allowPrivate: true }));
  const r2 = await open.search('x');
  assert.equal(r2.ok, true);
  mock.restoreAll();
});
