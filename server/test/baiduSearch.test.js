// 百度兼容层测试（统一 ImageResult 模型）
const { test, mock } = require('node:test');
const assert = require('node:assert');
const { searchImages } = require('../src/baiduSearch.js');

test('baidu search parses acjson into unified model', async () => {
  mock.method(global, 'fetch', async () => ({
    ok: true,
    json: async () => ({
      data: [
        { thumbURL: 'https://img0.baidu.com/it/u=1', middleURL: 'https://img0.baidu.com/it/m=1', fromPageTitle: '标题A', width: 800, height: 600, fromURL: 'https://p.example.com' },
        { thumbURL: 'http://insecure.example.com/t.jpg', middleURL: 'http://insecure.example.com/m.jpg', fromPageTitle: 'http图源保留' },
        { thumbURL: '', middleURL: '', fromPageTitle: '无图跳过' }
      ]
    })
  }));
  const images = await searchImages('测试', 10);
  assert.equal(images.length, 2, 'http 图源不再静默丢弃');
  assert.equal(images[0].imageUrl, 'https://img0.baidu.com/it/m=1');
  assert.equal(images[0].thumbnailUrl, 'https://img0.baidu.com/it/u=1');
  assert.equal(images[0].title, '标题A');
  assert.equal(images[0].width, 800);
  assert.equal(images[0].provider, 'baidu_page');
  mock.restoreAll();
});

test('baidu search handles fetch failure with empty result', async () => {
  mock.method(global, 'fetch', async () => { throw new Error('network'); });
  const images = await searchImages('x', 10);
  assert.deepEqual(images, []);
  mock.restoreAll();
});
