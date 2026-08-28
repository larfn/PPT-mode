// Image Provider 抽象基类 + 统一结果模型
//
// 统一 ImageResult：
//   { id, thumbnailUrl, imageUrl, width, height, source, sourceUrl, title,
//     author, license, mimeType, query, provider }
// 业务层（imageService / 路由 / 前端）只依赖这些统一字段，绝不触碰某个 provider 的页面结构。
//
// Provider 接口（至少实现 search / getMetadata / download）：
//   search(query, options)      → { ok, results: ImageResult[], error: {code,message}|null }
//   getMetadata(result)         → Promise<ImageResult>（可补充 license/author/mime 等）
//   download(result, opts)      → 默认委托安全下载器（downloadStore），业务层无需关心 provider 细节
//
// 错误约定：search 不 throw（参数错误除外）——网络/HTTP/解析失败都进 error 字段，
// 保证「页面结构解析失败时整个 API 不崩溃、返回 provider error、保留其他结果」。
'use strict';
const crypto = require('node:crypto');

class ProviderError extends Error {
  constructor(code, message) {
    super(message);
    this.code = code;
    this.name = 'ProviderError';
  }
}

// 统一结果 id：provider + 图片 URL 的短哈希（稳定、可去重、与页面结构无关）
function resultId(provider, imageUrl) {
  return provider + '-' + crypto.createHash('sha1').update(imageUrl || '').digest('hex').slice(0, 12);
}

// 把 provider 原始条目归一化为统一 ImageResult（兼容旧字段 fullUrl/pageUrl）
function normalizeResult(raw, ctx) {
  const provider = ctx.provider;
  const imageUrl = String(raw.imageUrl || raw.fullUrl || '').trim();
  const thumbnailUrl = String(raw.thumbnailUrl || imageUrl || '').trim();
  return {
    id: resultId(provider, imageUrl || thumbnailUrl),
    thumbnailUrl,
    imageUrl,
    width: Number(raw.width) || null,
    height: Number(raw.height) || null,
    source: raw.source || ctx.name || provider,
    sourceUrl: raw.sourceUrl || raw.pageUrl || null,
    title: String(raw.title || '').replace(/<[^>]*>/g, '').trim(),
    author: raw.author || null,
    license: raw.license || null,
    mimeType: raw.mimeType || null,
    query: ctx.query,
    provider
  };
}

class ImageProvider {
  constructor() {
    this.id = '';      // 唯一 provider id（如 'baidu_page'）
    this.name = '';    // 展示名（如 '百度图片'）
  }

  async search(query, options = {}) { // eslint-disable-line no-unused-vars
    throw new ProviderError('not-implemented', this.id + ' 未实现 search');
  }

  async getMetadata(result) {
    return result; // 默认不补充；未来 API provider 可覆写
  }

  async download(result, opts = {}) {
    // 默认下载：委托统一安全下载器（SSRF 校验/大小限制/MIME 校验/缓存都在下载器内）
    const { startDownload } = require('../downloadStore.js');
    const url = (result && (result.imageUrl || result.downloadUrl)) || '';
    if (!url) throw new ProviderError('bad-request', '缺少图片地址');
    return startDownload(url, { provider: this.id, source: result && result.title, ...opts });
  }
}

module.exports = { ImageProvider, ProviderError, normalizeResult, resultId };
