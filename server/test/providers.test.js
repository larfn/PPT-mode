// Provider 架构单元测试：统一模型 / 解析隔离 / provider 错误不崩溃 / registry / 自定义源
// 隔离真实配置：APPDATA 指向临时目录，保证 listProviders/getProvider 结果确定
process.env.APPDATA = require('node:os').tmpdir();

const { test, mock } = require('node:test');
const assert = require('node:assert');
const path = require('node:path');
const fs = require('node:fs');
const { BaiduPageProvider } = require('../src/providers/baiduPageProvider.js');
const { BingPageProvider } = require('../src/providers/bingPageProvider.js');
const { QihooPageProvider } = require('../src/providers/qihooPageProvider.js');
const { getProvider, listProviders, listBuiltins } = require('../src/providers/index.js');
const { searchImages } = require('../src/imageService.js');
const { searchImages: searchBaiduCompat } = require('../src/baiduSearch.js');
const { searchImages: searchBingCompat } = require('../src/bingSearch.js');
const { CONFIG_FILE, saveConfig, defaultConfig } = require('../src/config.js');
const { _setDnsLookup } = require('../src/downloadSecurity.js');

// 测试环境真实 DNS 不可靠：注入公网地址
_setDnsLookup(async () => [{ address: '93.184.216.34', family: 4 }]);

test('registry: builtin providers registered; unknown returns null', () => {
  assert.ok(getProvider('baidu_page'));
  assert.ok(getProvider('bing_page'));
  assert.ok(getProvider('qihoo_page'));
  assert.equal(getProvider('nope'), null);
  const builtinIds = listBuiltins().map((p) => p.id);
  assert.deepEqual(builtinIds.sort(), ['baidu_page', 'bing_page', 'qihoo_page']);
  // 全量列表 = 内置 + 默认启用的自定义源（搜狗随默认配置启用）
  const ids = listProviders().map((p) => p.id);
  for (const id of ['baidu_page', 'bing_page', 'qihoo_page', 'sogou_page']) {
    assert.ok(ids.includes(id), id + ' 应在可用列表');
  }
});

test('registry: enabled custom source in config resolves to JsonApiProvider', () => {
  const cfg = defaultConfig();
  cfg.image.sources = [
    {
      id: 'my_src', name: '我的源', enabled: true,
      endpoint: 'https://api.example.com/search?q={query}&n={count}',
      resultsPath: 'list', fields: { imageUrl: 'url', title: 'title' }
    },
    { id: 'disabled_src', name: '已禁用', enabled: false, endpoint: 'https://api.example.com/x' }
  ];
  saveConfig(CONFIG_FILE, cfg);
  try {
    const p = getProvider('my_src');
    assert.ok(p, '启用源应可解析');
    assert.equal(p.id, 'my_src');
    assert.equal(getProvider('disabled_src'), null, '禁用源不可解析');
    const ids = listProviders().map((x) => x.id);
    assert.ok(ids.includes('my_src'));
    assert.ok(!ids.includes('disabled_src'));
  } finally {
    fs.rmSync(CONFIG_FILE, { force: true });
  }
});

test('baidu provider returns unified ImageResult with all fields', async () => {
  mock.method(global, 'fetch', async () => ({
    ok: true, status: 200,
    json: async () => ({
      data: [
        { thumbURL: 'https://img0.baidu.com/it/u=1', middleURL: 'https://img0.baidu.com/it/m=1', fromPageTitle: '标题A<b>标签</b>', width: 800, height: 600, fromURL: 'https://page.example.com/a' }
      ]
    })
  }));
  const r = await new BaiduPageProvider().search('测试', { count: 10, page: 1 });
  assert.equal(r.ok, true);
  assert.equal(r.error, null);
  assert.equal(r.results.length, 1);
  const img = r.results[0];
  assert.equal(img.imageUrl, 'https://img0.baidu.com/it/m=1');
  assert.equal(img.thumbnailUrl, 'https://img0.baidu.com/it/u=1');
  assert.equal(img.width, 800);
  assert.equal(img.height, 600);
  assert.equal(img.source, '百度图片');
  assert.equal(img.sourceUrl, 'https://page.example.com/a');
  assert.equal(img.title, '标题A标签', 'HTML 标签应被剥离');
  assert.equal(img.query, '测试');
  assert.equal(img.provider, 'baidu_page');
  assert.ok(img.id.startsWith('baidu_page-'));
  assert.equal(img.author, null);
  assert.equal(img.license, null);
  assert.equal(img.mimeType, null);
  mock.restoreAll();
});

test('baidu provider: http image urls are kept (download layer enforces safety)', async () => {
  mock.method(global, 'fetch', async () => ({
    ok: true, status: 200,
    json: async () => ({ data: [{ thumbURL: 'https://t.example.com/t.jpg', middleURL: 'http://img.example.com/m.jpg', fromPageTitle: 'http图源' }] })
  }));
  const r = await new BaiduPageProvider().search('x', { count: 5 });
  assert.equal(r.results.length, 1, 'http 图源不再被静默丢弃（下载安全由下载层负责）');
  assert.equal(r.results[0].imageUrl, 'http://img.example.com/m.jpg');
  mock.restoreAll();
});

test('baidu provider: network failure → provider error, not throw', async () => {
  mock.method(global, 'fetch', async () => { throw new Error('network down'); });
  const r = await new BaiduPageProvider().search('x');
  assert.equal(r.ok, false);
  assert.equal(r.results.length, 0);
  assert.equal(r.error.code, 'network');
  mock.restoreAll();
});

test('baidu provider: malformed json → parse error, not throw', async () => {
  mock.method(global, 'fetch', async () => ({ ok: true, status: 200, json: async () => { throw new Error('Unexpected token'); } }));
  const r = await new BaiduPageProvider().search('x');
  assert.equal(r.ok, false);
  assert.equal(r.error.code, 'parse');
  mock.restoreAll();
});

test('baidu provider: changed structure (no data) → parse error with empty results', async () => {
  mock.method(global, 'fetch', async () => ({ ok: true, status: 200, json: async () => ({ somethingElse: 1 }) }));
  const r = await new BaiduPageProvider().search('x');
  assert.equal(r.ok, true);
  assert.equal(r.results.length, 0);
  assert.equal(r.error.code, 'parse');
  mock.restoreAll();
});

test('bing provider parses m= JSON into unified model; bad entries skipped', async () => {
  const html = '<div class="iusc" m="{&quot;murl&quot;:&quot;https://cdn.example.com/a.jpg&quot;,&quot;turl&quot;:&quot;https://cdn.example.com/thumb.jpg&quot;,&quot;purl&quot;:&quot;https://page.example.com/a&quot;,&quot;t&quot;:&quot;A test image&quot;,&quot;mw&quot;:800,&quot;mh&quot;:600}"></div>'
    + '<div class="iusc" m="not json"></div>'
    + '<div class="iusc" m="{&quot;turl&quot;:&quot;https://x/t.jpg&quot;}"></div>';
  mock.method(global, 'fetch', async () => ({ ok: true, status: 200, text: async () => html }));
  const r = await new BingPageProvider().search('test', { count: 10 });
  assert.equal(r.ok, true);
  assert.equal(r.results.length, 1, '坏条目跳过、好条目保留');
  const img = r.results[0];
  assert.equal(img.imageUrl, 'https://cdn.example.com/a.jpg');
  assert.equal(img.thumbnailUrl, 'https://cdn.example.com/thumb.jpg');
  assert.equal(img.width, 800);
  assert.equal(img.sourceUrl, 'https://page.example.com/a');
  assert.equal(img.provider, 'bing_page');
  mock.restoreAll();
});

test('bing provider: no results → parse error, not throw', async () => {
  mock.method(global, 'fetch', async () => ({ ok: true, status: 200, text: async () => '<html><body>no images</body></html>' }));
  const r = await new BingPageProvider().search('x');
  assert.equal(r.ok, true);
  assert.equal(r.results.length, 0);
  assert.equal(r.error.code, 'parse');
  mock.restoreAll();
});

test('bing provider: timeout → timeout error', async () => {
  mock.method(global, 'fetch', async () => { const e = new Error('timed out'); e.name = 'TimeoutError'; throw e; });
  const r = await new BingPageProvider().search('x');
  assert.equal(r.error.code, 'timeout');
  mock.restoreAll();
});

test('qihoo provider parses list into unified model', async () => {
  mock.method(global, 'fetch', async (url) => {
    assert.ok(String(url).includes('image.so.com/j'), '应请求 360 接口');
    assert.ok(String(url).includes('q=cat'));
    return {
      ok: true, status: 200,
      json: async () => ({
        list: [
          { img: 'http://p0.so.qhimg.com/a.jpg', thumb: 'https://p0.ssl.qhimgs1.com/a.jpg', width: 896, height: 1152, title: 'cat', link: 'https://pic.360.com/detail/1', site: 'pic.360.com' },
          { img: 'http://p0.so.qhimg.com/b.jpg', thumb: 'https://p0.ssl.qhimgs1.com/b.jpg' }
        ]
      })
    };
  });
  const r = await new QihooPageProvider().search('cat', { count: 5 });
  assert.equal(r.ok, true);
  assert.equal(r.results.length, 2);
  const img = r.results[0];
  assert.equal(img.imageUrl, 'http://p0.so.qhimg.com/a.jpg');
  assert.equal(img.thumbnailUrl, 'https://p0.ssl.qhimgs1.com/a.jpg');
  assert.equal(img.width, 896);
  assert.equal(img.height, 1152);
  assert.equal(img.title, 'cat');
  assert.equal(img.sourceUrl, 'https://pic.360.com/detail/1');
  assert.equal(img.provider, 'qihoo_page');
  mock.restoreAll();
});

test('qihoo provider: network failure → provider error; empty list → parse error', async () => {
  mock.method(global, 'fetch', async () => { throw new Error('down'); });
  let r = await new QihooPageProvider().search('x');
  assert.equal(r.ok, false);
  assert.equal(r.error.code, 'network');
  mock.restoreAll();

  mock.method(global, 'fetch', async () => ({ ok: true, status: 200, json: async () => ({ list: [] }) }));
  r = await new QihooPageProvider().search('x');
  assert.equal(r.ok, true);
  assert.equal(r.results.length, 0);
  assert.equal(r.error.code, 'parse');
  mock.restoreAll();
});

test('compat layers return unified model arrays (old callers keep working)', async () => {
  mock.method(global, 'fetch', async (url) => {
    if (String(url).includes('image.baidu.com')) {
      return { ok: true, status: 200, json: async () => ({ data: [{ middleURL: 'https://img.example.com/m.jpg', thumbURL: 'https://img.example.com/t.jpg', width: 10 }] }) };
    }
    return { ok: true, status: 200, text: async () => '<div m="{&quot;murl&quot;:&quot;https://cdn.example.com/a.jpg&quot;,&quot;turl&quot;:&quot;https://cdn.example.com/t.jpg&quot;}"></div>' };
  });
  const bd = await searchBaiduCompat('x', 5);
  assert.equal(bd[0].imageUrl, 'https://img.example.com/m.jpg');
  const bg = await searchBingCompat('x', 5);
  assert.equal(bg[0].imageUrl, 'https://cdn.example.com/a.jpg');
  mock.restoreAll();
});

test('imageService: bing_api → not-implemented error; unknown provider throws', async () => {
  const r = await searchImages('bing_api', 'x', 5, 1);
  assert.equal(r.ok, false);
  assert.equal(r.error.code, 'not-implemented');
  await assert.rejects(() => searchImages('nope', 'x', 5, 1), /unknown provider/);
});

test('imageService: provider failure surfaces as error field (API does not crash)', async () => {
  mock.method(global, 'fetch', async () => { throw new Error('network'); });
  const r = await searchImages('baidu_page', 'x', 5, 1);
  assert.equal(r.ok, false);
  assert.equal(r.error.code, 'network');
  assert.equal(r.provider, 'baidu_page');
  mock.restoreAll();
});

test('imageService: custom source from config searchable; disabled not found', async () => {
  const cfg = defaultConfig();
  cfg.image.sources = [{
    id: 'custom_a', name: 'A源', enabled: true,
    endpoint: 'https://api.example.com/search?q={query}&per_page={count}',
    resultsPath: 'data.list',
    fields: { imageUrl: 'url', thumbnailUrl: 'thumb', width: 'w', height: 'h', title: 'title', sourceUrl: 'page' }
  }];
  saveConfig(CONFIG_FILE, cfg);
  try {
    mock.method(global, 'fetch', async (url) => {
      assert.ok(String(url).includes('q=%E7%8C%AB'), 'query 应被 URL 编码');
      assert.ok(String(url).includes('per_page=5'));
      return {
        ok: true, status: 200,
        json: async () => ({
          data: { list: [{ url: 'https://cdn.example.com/1.jpg', thumb: 'https://cdn.example.com/1t.jpg', w: 100, h: 200, title: '图1', page: 'https://page.example.com/1' }] }
        })
      };
    });
    const r = await searchImages('custom_a', '猫', 5, 1);
    assert.equal(r.ok, true);
    assert.equal(r.provider, 'custom_a');
    assert.equal(r.providerName, 'A源');
    assert.equal(r.results.length, 1);
    assert.equal(r.results[0].imageUrl, 'https://cdn.example.com/1.jpg');
    assert.equal(r.results[0].thumbnailUrl, 'https://cdn.example.com/1t.jpg');
    assert.equal(r.results[0].title, '图1');
    assert.equal(r.results[0].provider, 'custom_a');
  } finally {
    mock.restoreAll();
    fs.rmSync(CONFIG_FILE, { force: true });
  }
});