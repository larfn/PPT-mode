// 360 图片 Provider（兼容实现）：解析 image.so.com 公开 JSON 接口（免 Key、国内直连）
'use strict';
const { ImageProvider, normalizeResult } = require('./providerBase.js');

const SO_URL = 'https://image.so.com/j';
const UA = 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0 Safari/537.36';
const TIMEOUT_MS = 15000;

class QihooPageProvider extends ImageProvider {
  constructor() {
    super();
    this.id = 'qihoo_page';
    this.name = '360 图片';
  }

  async search(query, options = {}) {
    const count = Math.min(Number(options.count) || 30, 60);
    const page = Math.max(1, Number(options.page) || 1);
    const params = new URLSearchParams({ q: query, pn: String(page), ps: String(count) });
    const url = SO_URL + '?' + params.toString();

    let payload = null;
    try {
      const res = await fetch(url, {
        headers: { 'User-Agent': UA, 'Referer': 'https://image.so.com/' },
        signal: AbortSignal.timeout(TIMEOUT_MS)
      });
      if (!res.ok) {
        return { ok: false, results: [], error: { code: 'http', message: '360 图片响应异常（HTTP ' + res.status + '）' } };
      }
      try {
        payload = await res.json();
      } catch {
        return { ok: false, results: [], error: { code: 'parse', message: '360 图片响应解析失败（结构可能已变化）' } };
      }
    } catch (e) {
      if (e && e.name === 'TimeoutError') {
        return { ok: false, results: [], error: { code: 'timeout', message: '360 图片请求超时' } };
      }
      return { ok: false, results: [], error: { code: 'network', message: '360 图片请求失败：' + ((e && e.message) || String(e)) } };
    }

    const list = (payload && Array.isArray(payload.list)) ? payload.list : [];
    if (!list.length) {
      return { ok: true, results: [], error: { code: 'parse', message: '360 图片未返回可用结果（页面结构可能已变化）' } };
    }
    const results = [];
    for (const d of list) {
      if (!d || typeof d !== 'object') continue;
      const imageUrl = d.img || d.thumb || '';
      const thumbnailUrl = d.thumb || d.img || '';
      if (!imageUrl) continue;
      results.push(normalizeResult({
        imageUrl,
        thumbnailUrl,
        width: d.width,
        height: d.height,
        title: d.title || '',
        sourceUrl: d.link || null,
        source: d.site || this.name
      }, { provider: this.id, name: this.name, query }));
      if (results.length >= count) break;
    }
    if (!results.length) {
      return { ok: true, results: [], error: { code: 'parse', message: '360 图片未解析到结果（页面结构可能已变化）' } };
    }
    return { ok: true, results, error: null };
  }
}

// 旧接口兼容：返回统一模型数组
async function searchImagesCompat(query, count, page) {
  const r = await new QihooPageProvider().search(query, { count, page });
  return r.results;
}

module.exports = { QihooPageProvider, searchImages: searchImagesCompat };
