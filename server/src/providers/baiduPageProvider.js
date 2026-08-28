// 百度图片 Provider（兼容实现）：解析公开 JSON 接口 image.baidu.com/search/acjson
'use strict';
const { ImageProvider, normalizeResult } = require('./providerBase.js');

const BAIDU_ACJSON = 'https://image.baidu.com/search/acjson';
const UA = 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0 Safari/537.36';
const TIMEOUT_MS = 15000;

class BaiduPageProvider extends ImageProvider {
  constructor() {
    super();
    this.id = 'baidu_page';
    this.name = '百度图片';
  }

  async search(query, options = {}) {
    const count = Math.min(Number(options.count) || 30, 60);
    const page = Math.max(1, Number(options.page) || 1);
    const pn = (page - 1) * count;
    const params = new URLSearchParams({
      tn: 'resultjson_com', word: query, queryWord: query,
      pn: String(pn), rn: String(count), ie: 'utf-8', oe: 'utf-8', cl: '2', lm: '-1'
    });
    const url = BAIDU_ACJSON + '?' + params.toString();

    let payload = null;
    try {
      const res = await fetch(url, {
        headers: { 'User-Agent': UA, 'Referer': 'https://image.baidu.com/' },
        signal: AbortSignal.timeout(TIMEOUT_MS)
      });
      if (!res.ok) {
        return { ok: false, results: [], error: { code: 'http', message: '百度图片响应异常（HTTP ' + res.status + '）' } };
      }
      try {
        payload = await res.json();
      } catch {
        return { ok: false, results: [], error: { code: 'parse', message: '百度图片响应解析失败（结构可能已变化）' } };
      }
    } catch (e) {
      if (e && e.name === 'TimeoutError') {
        return { ok: false, results: [], error: { code: 'timeout', message: '百度图片请求超时' } };
      }
      return { ok: false, results: [], error: { code: 'network', message: '百度图片请求失败：' + ((e && e.message) || String(e)) } };
    }

    // 结构解析：字段缺失/JSON 变化 → 返回空结果 + 解析错误，绝不 throw
    const list = (payload && Array.isArray(payload.data)) ? payload.data : [];
    if (!list.length) {
      return { ok: true, results: [], error: { code: 'parse', message: '百度图片未返回可用结果（页面结构可能已变化）' } };
    }
    const results = [];
    for (const d of list) {
      if (!d || typeof d !== 'object') continue;
      const imageUrl = d.middleURL || d.thumbURL || '';
      const thumbnailUrl = d.thumbURL || d.middleURL || '';
      if (!imageUrl) continue;
      results.push(normalizeResult({
        imageUrl,
        thumbnailUrl,
        pageUrl: d.fromURL || '',
        width: d.width,
        height: d.height,
        title: String(d.fromPageTitle || d.fromPageTitleTag || '').replace(/<[^>]*>/g, '').trim(),
        source: this.name
      }, { provider: this.id, name: this.name, query }));
      if (results.length >= count) break;
    }
    return { ok: true, results, error: null };
  }
}

// 旧接口兼容：返回统一模型数组（历史调用方/测试使用）
async function searchImagesCompat(query, count, page) {
  const r = await new BaiduPageProvider().search(query, { count, page });
  return r.results;
}

module.exports = { BaiduPageProvider, searchImages: searchImagesCompat };
