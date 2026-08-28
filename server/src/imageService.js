// 图片服务：通过 Provider 注册表分发搜索，业务层只依赖统一 ImageResult。
//  - provider 内部错误（网络/超时/解析失败）返回 { ok:false, error }，不会让 API 崩溃
//  - 未实现的 provider（如 bing_api）返回明确的 not-implemented 错误
//  - 下载统一走 downloadStore（SSRF/MIME/大小/缓存），provider 不关心
//  - 性能：搜索结果服务端缓存（key=provider|query|count|page，TTL 10 分钟，容量 50）
//    重复搜索（翻页/回退）秒回，首屏 <3s 更稳。
'use strict';
const { getProvider, listProviders } = require('./providers/index.js');

const CACHE_TTL_MS = 10 * 60 * 1000; // 10 分钟
const MAX_CACHE = 50;
const searchCache = new Map();

function cacheKey(providerId, query, count, page) {
  return providerId + '|' + query + '|' + count + '|' + page;
}

async function searchImages(providerId, query, count, page) {
  const provider = getProvider(providerId);
  if (!provider) {
    if (providerId === 'bing_api') {
      return { ok: false, results: [], error: { code: 'not-implemented', message: 'Bing API（Azure）尚未实现，请改用「搜索引擎搜图」来源' }, provider: providerId, providerName: null };
    }
    throw new Error('unknown provider: ' + providerId);
  }
  const key = cacheKey(providerId, String(query || ''), count, page);
  const hit = searchCache.get(key);
  if (hit && Date.now() - hit.at < CACHE_TTL_MS) {
    return { ok: hit.ok, results: hit.results, error: hit.error, provider: providerId, providerName: provider.name, fromCache: true };
  }
  const r = await provider.search(query, { count, page });
  const out = { ok: r.ok, results: r.results, error: r.error, provider: providerId, providerName: provider.name };
  if (r.ok && r.results && r.results.length) {
    // 深拷贝入缓存：防止调用方修改结果对象污染缓存（结果全是纯 JSON，拷贝开销可忽略）
    searchCache.set(key, { at: Date.now(), ok: r.ok, results: JSON.parse(JSON.stringify(r.results)), error: r.error });
    if (searchCache.size > MAX_CACHE) {
      const k = searchCache.keys().next().value;
      if (k) searchCache.delete(k);
    }
  }
  return out;
}

module.exports = { searchImages, listProviders, _clearSearchCache: () => searchCache.clear() };