// 必应兼容层测试（统一 ImageResult 模型）
const { test, mock } = require('node:test');
const assert = require('node:assert');
const { searchImages } = require('../src/bingSearch.js');

test('searchImages parses m= JSON from bing html into unified model', async () => {
  const html = '<div class="iusc" m="{&quot;murl&quot;:&quot;https://cdn.example.com/a.jpg&quot;,&quot;turl&quot;:&quot;https://cdn.example.com/thumb.jpg&quot;,&quot;purl&quot;:&quot;https://page.example.com/a&quot;,&quot;t&quot;:&quot;A test image&quot;,&quot;mw&quot;:800,&quot;mh&quot;:600}"></div>';
  mock.method(global, 'fetch', async () => ({ ok: true, text: async () => html }));
  const images = await searchImages('test', 2);
  assert.equal(images.length, 1);
  assert.equal(images[0].imageUrl, 'https://cdn.example.com/a.jpg');
  assert.equal(images[0].thumbnailUrl, 'https://cdn.example.com/thumb.jpg');
  assert.equal(images[0].width, 800);
  assert.equal(images[0].provider, 'bing_page');
  mock.restoreAll();
});

test('searchImages handles fetch failure with empty result', async () => {
  mock.method(global, 'fetch', async () => { throw new Error('network'); });
  const images = await searchImages('test', 2);
  assert.deepEqual(images, []);
  mock.restoreAll();
});
