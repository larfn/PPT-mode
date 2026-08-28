// Bing 图片 Provider（兼容实现）：解析 www.bing.com/images/search 页面内嵌 m="..." JSON
'use strict';
const { ImageProvider, normalizeResult } = require('./providerBase.js');

const BING_URL = 'https://www.bing.com/images/search';
const UA = 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0 Safari/537.36';
const TIMEOUT_MS = 15000;

// HTML 实体解码（含数字实体 &#xx; 与 &#xXX;，覆盖页面改版常见情况）
function decodeEntities(s) {
  return s
    .replace(/&quot;/g, '"')
    .replace(/&amp;/g, '&')
    .replace(/&#39;|&apos;/g, "'")
    .replace(/&lt;/g, '<')
    .replace(/&gt;/g, '>')
    .replace(/&#x([0-9a-fA-F]+);/g, (_, h) => String.fromCodePoint(parseInt(h, 16)))
    .replace(/&#(\d+);/g, (_, d) => String.fromCodePoint(parseInt(d, 10)));
}

class BingPageProvider extends ImageProvider {
  constructor() {
    super();
    this.id = 'bing_page';
    this.name = '必应图片';
  }

  async search(query, options = {}) {
    const count = Math.min(Number(options.count) || 9, 35);
    const page = Math.max(1, Number(options.page) || 1);
    const first = (page - 1) * count + 1;
    const params = new URLSearchParams({ q: query, count: String(count), first: String(first), mkt: 'zh-CN' });

    let html = '';
    try {
      const res = await fetch(BING_URL + '?' + params.toString(), {
        headers: { 'User-Agent': UA },
        signal: AbortSignal.timeout(TIMEOUT_MS)
      });
      if (!res.ok) {
        return { ok: false, results: [], error: { code: 'http', message: '必应图片响应异常（HTTP ' + res.status + '）' } };
      }
      html = await res.text();
    } catch (e) {
      if (e && e.name === 'TimeoutError') {
        return { ok: false, results: [], error: { code: 'timeout', message: '必应图片请求超时' } };
      }
      return { ok: false, results: [], error: { code: 'network', message: '必应图片请求失败：' + ((e && e.message) || String(e)) } };
    }

    // 页面解析：逐个条目独立 try/catch，坏条目跳过、保留好条目；整体无匹配 → 解析错误
    const results = [];
    const regex = /m="([^"]+)"/g;
    let match;
    while ((match = regex.exec(html)) !== null && results.length < count) {
      try {
        const data = JSON.parse(decodeEntities(match[1]));
        if (!data || typeof data !== 'object' || !data.murl) continue;
        results.push(normalizeResult({
          imageUrl: data.murl,
          thumbnailUrl: data.turl || '',
          pageUrl: data.purl || '',
          width: data.mw,
          height: data.mh,
          title: data.t || '',
          source: this.name
        }, { provider: this.id, name: this.name, query }));
      } catch {
        // 跳过无法解析的条目（页面改版时保持其他结果可用）
      }
    }
    if (!results.length) {
      return { ok: true, results: [], error: { code: 'parse', message: '必应图片未解析到结果（页面结构可能已变化）' } };
    }
    return { ok: true, results, error: null };
  }
}

// 旧接口兼容：返回统一模型数组
async function searchImagesCompat(query, count, page) {
  const r = await new BingPageProvider().search(query, { count, page });
  return r.results;
}

module.exports = { BingPageProvider, searchImages: searchImagesCompat };
