// 下载系统测试：安全校验 / SSRF / 缓存 / 重定向 / 超限 / 非图片 / SVG / 空文件
const { test, mock } = require('node:test');
const assert = require('node:assert');
const fs = require('node:fs');
const os = require('node:os');
const path = require('node:path');
const { createApp } = require('../src/index.js');
const { _setDnsLookup } = require('../src/downloadSecurity.js');
const { _clearCache, _fetchImageWithGuard } = require('../src/downloadStore.js');

const DL_DIR = path.join(os.tmpdir(), 'ppt-dl-test-' + Date.now());
process.env.PPT_DOWNLOAD_DIR = DL_DIR;

// 公网 DNS 注入（测试环境真实 DNS 不可靠）
_setDnsLookup(async () => [{ address: '93.184.216.34', family: 4 }]);

test('GET /api/images/providers lists available providers', async () => {
  const app = createApp();
  const server = await listen(app);
  try {
    const base = 'http://127.0.0.1:' + server.address().port;
    const body = await (await fetch(base + '/api/images/providers')).json();
    const ids = body.providers.map((p) => p.id);
    assert.ok(ids.includes('baidu_page') && ids.includes('bing_page'));
  } finally { server.close(); }
});

test('search API: provider error surfaces as error field, page/hasMore intact', async () => {
  const realFetch = global.fetch.bind(global);
  mock.method(global, 'fetch', async () => { throw new Error('network'); });
  const app = createApp();
  const server = await listen(app);
  try {
    const base = 'http://127.0.0.1:' + server.address().port;
    const res = await realFetch(base + '/api/images/search', {
      method: 'POST', headers: { 'content-type': 'application/json' },
      body: JSON.stringify({ query: 'test', count: 9, page: 2, provider: 'baidu_page' })
    });
    assert.equal(res.status, 200, 'provider 内部错误不崩溃');
    const body = await res.json();
    assert.deepEqual(body.images, []);
    assert.equal(body.page, 2);
    assert.equal(body.error.code, 'network');
    assert.equal(body.provider, 'baidu_page');
  } finally { server.close(); mock.restoreAll(); }
});

test('search API: bing_api returns not-implemented error', async () => {
  const app = createApp();
  const server = await listen(app);
  try {
    const base = 'http://127.0.0.1:' + server.address().port;
    const res = await fetch(base + '/api/images/search', {
      method: 'POST', headers: { 'content-type': 'application/json' },
      body: JSON.stringify({ query: 'x', provider: 'bing_api' })
    });
    const body = await res.json();
    assert.equal(body.error.code, 'not-implemented');
  } finally { server.close(); }
});

test('search API: unknown provider → 400', async () => {
  const app = createApp();
  const server = await listen(app);
  try {
    const base = 'http://127.0.0.1:' + server.address().port;
    const res = await fetch(base + '/api/images/search', {
      method: 'POST', headers: { 'content-type': 'application/json' },
      body: JSON.stringify({ query: 'x', provider: 'nope' })
    });
    assert.equal(res.status, 400);
  } finally { server.close(); }
});



const PNG_1PX = Buffer.from('iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8BQDwAEhQGAhKmMIQAAAABJRU5ErkJggg==', 'base64');

function listen(app) {
  return new Promise((resolve) => {
    const server = app.listen(0, '127.0.0.1', () => resolve(server));
  });
}
async function pollTask(base, taskId, fetcher = global.fetch) {
  let body = {};
  for (let i = 0; i < 50 && !body.done; i++) {
    const sres = await fetcher(base + '/api/images/download/' + taskId);
    body = await sres.json();
    if (!body.done) await new Promise((r) => setTimeout(r, 15));
  }
  return body;
}

test('download rejects missing url', async () => {
  const app = createApp();
  const server = await listen(app);
  try {
    const base = 'http://127.0.0.1:' + server.address().port;
    const res = await fetch(base + '/api/images/download', {
      method: 'POST', headers: { 'content-type': 'application/json' }, body: JSON.stringify({})
    });
    assert.equal(res.status, 400);
  } finally { server.close(); }
});

test('SSRF: localhost / 127.0.0.1 / file protocol are rejected in the task', async () => {
  const app = createApp();
  const server = await listen(app);
  try {
    const base = 'http://127.0.0.1:' + server.address().port;
    for (const url of ['https://localhost/x.jpg', 'https://127.0.0.1/x.jpg', 'http://10.0.0.1/x.jpg', 'file:///etc/passwd', 'ftp://x.com/a.jpg']) {
      const res = await fetch(base + '/api/images/download', {
        method: 'POST', headers: { 'content-type': 'application/json' }, body: JSON.stringify({ url })
      });
      assert.equal(res.status, 200);
      const { taskId } = await res.json();
      const body = await pollTask(base, taskId);
      assert.equal(body.done, false, url + ' 不应下载成功');
      assert.ok(body.error && body.error.length > 0, url + ' 应有错误信息: ' + body.error);
    }
  } finally { server.close(); }
});

test('SSRF: domain resolving to private IP is rejected', async () => {
  _setDnsLookup(async () => [{ address: '192.168.1.10', family: 4 }]);
  const app = createApp();
  const server = await listen(app);
  try {
    const base = 'http://127.0.0.1:' + server.address().port;
    const res = await fetch(base + '/api/images/download', {
      method: 'POST', headers: { 'content-type': 'application/json' },
      body: JSON.stringify({ url: 'https://internal.example.com/a.jpg' })
    });
    const { taskId } = await res.json();
    const body = await pollTask(base, taskId);
    assert.ok(body.error.includes('内网'), body.error);
  } finally { server.close(); }
  _setDnsLookup(async () => [{ address: '93.184.216.34', family: 4 }]);
});

test('download success: content-type + magic check + atomic file + dataUrl', async () => {
  const realFetch = global.fetch.bind(global);
  mock.method(global, 'fetch', async (url, opts) => {
    assert.equal(opts.redirect, 'manual', '重定向必须手动跟随以便逐跳校验');
    return {
      ok: true, status: 200,
      headers: new Map([['content-type', 'image/png']]),
      body: null,
      arrayBuffer: async () => PNG_1PX.buffer.slice(PNG_1PX.byteOffset, PNG_1PX.byteOffset + PNG_1PX.byteLength)
    };
  });
  const app = createApp();
  const server = await listen(app);
  try {
    const base = 'http://127.0.0.1:' + server.address().port;
    const res = await realFetch(base + '/api/images/download', {
      method: 'POST', headers: { 'content-type': 'application/json' },
      body: JSON.stringify({ url: 'https://cdn.example.com/a.png', provider: 'baidu_page' })
    });
    assert.equal(res.status, 200);
    const { taskId } = await res.json();
    const body = await pollTask(base, taskId, realFetch);
    assert.equal(body.done, true);
    assert.ok(body.dataUrl.startsWith('data:image/png;base64,'));
    assert.ok(body.fileName.endsWith('.png'));
    assert.ok(fs.existsSync(body.filePath));
    assert.equal(body.provider, 'baidu_page');
  } finally { server.close(); mock.restoreAll(); }
});

test('download cache: same provider+url downloads only once (fromCache)', async () => {
  _clearCache();
  const realFetch = global.fetch.bind(global);
  let fetchCount = 0;
  mock.method(global, 'fetch', async () => {
    fetchCount++;
    return {
      ok: true, status: 200,
      headers: new Map([['content-type', 'image/png']]),
      body: null,
      arrayBuffer: async () => PNG_1PX.buffer.slice(PNG_1PX.byteOffset, PNG_1PX.byteOffset + PNG_1PX.byteLength)
    };
  });
  const app = createApp();
  const server = await listen(app);
  try {
    const base = 'http://127.0.0.1:' + server.address().port;
    for (let i = 0; i < 2; i++) {
      const res = await realFetch(base + '/api/images/download', {
        method: 'POST', headers: { 'content-type': 'application/json' },
        body: JSON.stringify({ url: 'https://cdn.example.com/cached.png', provider: 'baidu_page' })
      });
      const { taskId } = await res.json();
      const body = await pollTask(base, taskId, realFetch);
      assert.equal(body.done, true);
      if (i === 1) assert.equal(body.fromCache, true, '第二次应命中缓存');
    }
    assert.equal(fetchCount, 1, '网络请求应只发生一次');
  } finally { server.close(); mock.restoreAll(); _clearCache(); }
});

test('redirect: follows with per-hop SSRF re-check, up to limit', async () => {
  let hops = 0;
  mock.method(global, 'fetch', async (url) => {
    hops++;
    if (hops === 1) {
      return { ok: false, status: 302, headers: new Map([['location', 'https://cdn.example.com/final.png']]), body: null, arrayBuffer: async () => new ArrayBuffer(0) };
    }
    return {
      ok: true, status: 200,
      headers: new Map([['content-type', 'image/png']]),
      body: null,
      arrayBuffer: async () => PNG_1PX.buffer.slice(PNG_1PX.byteOffset, PNG_1PX.byteOffset + PNG_1PX.byteLength)
    };
  });
  const r = await _fetchImageWithGuard('https://example.com/redirect.png');
  assert.equal(r.mime, 'image/png');
  assert.equal(hops, 2);
  mock.restoreAll();
});

test('redirect to private IP is rejected on the second hop', async () => {
  let hops = 0;
  mock.method(global, 'fetch', async () => {
    hops++;
    return { ok: false, status: 302, headers: new Map([['location', 'https://127.0.0.1/steal.png']]), body: null, arrayBuffer: async () => new ArrayBuffer(0) };
  });
  await assert.rejects(() => _fetchImageWithGuard('https://example.com/r.png'), /本机|内网|仅支持/);
  mock.restoreAll();
});

test('non-image response / svg / empty body / wrong magic are rejected', async () => {
  const cases = [
    { ct: 'text/html', body: Buffer.from('<html>'), expect: /非图片响应/ },
    { ct: 'image/svg+xml', body: Buffer.from('<svg></svg>'), expect: /SVG/ },
    { ct: 'image/png', body: Buffer.alloc(0), expect: /为空/ },
    { ct: 'image/png', body: Buffer.from('NOTANIMAGE', 'utf8'), expect: /不支持的图片格式|内容与声明格式不符/ }
  ];
  for (const c of cases) {
    mock.method(global, 'fetch', async () => ({
      ok: true, status: 200,
      headers: new Map([['content-type', c.ct]]),
      body: null,
      arrayBuffer: async () => c.body.buffer.slice(c.body.byteOffset, c.body.byteOffset + c.body.byteLength)
    }));
    await assert.rejects(() => _fetchImageWithGuard('https://cdn.example.com/x'), c.expect, c.ct);
    mock.restoreAll();
  }
});

test('oversized image is aborted (memory protection)', async () => {
  const big = Buffer.alloc(21 * 1024 * 1024, 0x89); // > 20MB 上限，且前 4 字节 0x89... 会被 size 检查拦截
  mock.method(global, 'fetch', async () => ({
    ok: true, status: 200,
    headers: new Map([['content-type', 'image/png']]),
    body: null,
    arrayBuffer: async () => big.buffer.slice(big.byteOffset, big.byteOffset + big.byteLength)
  }));
  await assert.rejects(() => _fetchImageWithGuard('https://cdn.example.com/huge.png'), /过大/);
  mock.restoreAll();
});

test('GET /api/images/providers lists available providers', async () => {
  const app = createApp();
  const server = await listen(app);
  try {
    const base = 'http://127.0.0.1:' + server.address().port;
    const body = await (await fetch(base + '/api/images/providers')).json();
    const ids = body.providers.map((p) => p.id);
    assert.ok(ids.includes('baidu_page') && ids.includes('bing_page'));
  } finally { server.close(); }
});
