// 通用 JSON API Provider：把「自定义源定义」变成可搜索的 ImageProvider。
// 定义支持：
//   endpoint 模板 {query} {count} {page} {start} {key}（URL 编码；headers 值中 {key} 原样替换）
//   headers / cookies（对象；cookies 自动拼 Cookie 头）
//   resultsPath 结果数组路径（留空 = 整个 JSON 即数组）
//   fields 字段映射（点分路径 + 数组下标，如 'data.items'、'imageinfo[0].url'）
// 错误约定与内置 provider 一致：search 不 throw（参数校验除外），错误进 error 字段。
'use strict';
const { ImageProvider, ProviderError, normalizeResult } = require('./providerBase.js');
const { validateDownloadUrl } = require('../downloadSecurity.js');

const UA = 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0 Safari/537.36';
const TIMEOUT_MS = 15000;

const FIELD_KEYS = ['imageUrl', 'thumbnailUrl', 'width', 'height', 'source', 'sourceUrl', 'title', 'author', 'license', 'mimeType'];
const ID_RE = /^[a-zA-Z0-9_-]{1,64}$/;

// ---------- 路径取值：a.b 与 a[0].b 组合 ----------
function getPath(obj, pathStr) {
  if (pathStr == null || pathStr === '') return obj;
  const parts = [];
  const re = /([^.\[\]]+)|\[(\d+)\]/g;
  let m;
  while ((m = re.exec(pathStr)) !== null) {
    parts.push(m[1] !== undefined ? m[1] : Number(m[2]));
  }
  let cur = obj;
  for (const p of parts) {
    if (cur == null) return undefined;
    cur = cur[p];
  }
  return cur;
}

// ---------- 模板替换 ----------
function fillTemplate(tpl, vars, { encode = true } = {}) {
  return String(tpl).replace(/\{(\w+)\}/g, (_, k) => {
    if (!(k in vars)) return '{' + k + '}';
    const v = vars[k];
    return encode ? encodeURIComponent(String(v)) : String(v);
  });
}

// ---------- 源定义校验/清洗 ----------
function validateSourceDef(src) {
  const errors = [];
  if (!src || typeof src !== 'object') return { ok: false, errors: ['源定义必须是 JSON 对象'] };
  if (typeof src.id !== 'string' || !ID_RE.test(src.id)) errors.push('id 必填：字母/数字/下划线/中划线，1-64 位');
  if (typeof src.name !== 'string' || !src.name.trim()) errors.push('name 必填');
  if (typeof src.endpoint !== 'string' || !src.endpoint.trim()) {
    errors.push('endpoint 必填（接口地址模板）');
  } else {
    let parsed;
    try { parsed = new URL(src.endpoint); } catch { parsed = null; }
    if (!parsed || (parsed.protocol !== 'http:' && parsed.protocol !== 'https:')) {
      errors.push('endpoint 必须是 http:// 或 https:// 开头的合法 URL');
    }
  }
  if (src.headers != null && (typeof src.headers !== 'object' || Array.isArray(src.headers))) errors.push('headers 必须是对象');
  if (src.cookies != null && (typeof src.cookies !== 'object' || Array.isArray(src.cookies))) errors.push('cookies 必须是对象');
  if (src.resultsPath != null && typeof src.resultsPath !== 'string') errors.push('resultsPath 必须是字符串');
  if (src.fields != null && (typeof src.fields !== 'object' || Array.isArray(src.fields))) errors.push('fields 必须是对象');
  if (src.key != null && typeof src.key !== 'string') errors.push('key 必须是字符串');
  if (src.keyRequired === true && (typeof src.key !== 'string' || !src.key.trim() || src.key.includes('****'))) errors.push('该图源需要 API Key');
  if (src.enabled != null && typeof src.enabled !== 'boolean') errors.push('enabled 必须是布尔值');
  if (src.allowPrivate != null && typeof src.allowPrivate !== 'boolean') errors.push('allowPrivate 必须是布尔值');
  return errors.length ? { ok: false, errors } : { ok: true, errors: [] };
}

function sanitizeSourceDef(src) {
  const fields = {};
  if (src.fields && typeof src.fields === 'object') {
    for (const k of FIELD_KEYS) {
      const v = src.fields[k];
      if (typeof v === 'string') fields[k] = v;
    }
  }
  const headers = {};
  if (src.headers && typeof src.headers === 'object' && !Array.isArray(src.headers)) {
    for (const [k, v] of Object.entries(src.headers)) {
      if (typeof k === 'string' && typeof v === 'string') headers[k] = v;
    }
  }
  const cookies = {};
  if (src.cookies && typeof src.cookies === 'object' && !Array.isArray(src.cookies)) {
    for (const [k, v] of Object.entries(src.cookies)) {
      if (typeof k === 'string' && typeof v === 'string') cookies[k] = v;
    }
  }
  return {
    id: String(src.id || '').trim(),
    name: String(src.name || '').trim(),
    endpoint: String(src.endpoint || '').trim(),
    headers,
    cookies,
    resultsPath: typeof src.resultsPath === 'string' ? src.resultsPath : '',
    fields,
    key: typeof src.key === 'string' ? src.key : '',
    enabled: src.enabled !== false,
    allowPrivate: src.allowPrivate === true,
    note: typeof src.note === 'string' ? src.note : '',
    keyRequired: src.keyRequired === true
  };
}

// ---------- SSRF：自定义源接口地址不允许指向内网/本机（allowPrivate=true 显式放行） ----------
const hostCache = new Map(); // host -> true（校验通过）
async function ensureSafeEndpoint(endpoint, allowPrivate) {
  let parsed;
  try { parsed = new URL(endpoint); } catch { throw new ProviderError('bad-request', '自定义源接口地址不是合法 URL'); }
  if (parsed.protocol !== 'http:' && parsed.protocol !== 'https:') {
    throw new ProviderError('bad-request', '自定义源接口仅支持 http/https');
  }
  if (allowPrivate) return;
  const host = parsed.hostname.toLowerCase();
  if (hostCache.has(host)) return;
  const v = await validateDownloadUrl(endpoint);
  if (!v.ok) throw new ProviderError('bad-request', '自定义源地址未通过安全校验：' + v.error);
  hostCache.set(host, true);
}

class JsonApiProvider extends ImageProvider {
  constructor(def) {
    super();
    this.def = def;
    this.id = def.id;
    this.name = def.name;
  }

  async search(query, options = {}) {
    const count = Math.min(Number(options.count) || 9, 60);
    const page = Math.max(1, Number(options.page) || 1);
    const def = this.def;

    try {
      await ensureSafeEndpoint(def.endpoint, def.allowPrivate === true);
    } catch (e) {
      return { ok: false, results: [], error: { code: 'bad-request', message: e.message } };
    }

    const vars = {
      query: String(query || ''),
      count,
      page,
      start: (page - 1) * count,
      key: def.key || ''
    };
    const url = fillTemplate(def.endpoint, vars);

    const headers = { 'User-Agent': UA };
    for (const [k, v] of Object.entries(def.headers || {})) {
      headers[k] = fillTemplate(v, { key: def.key || '' }, { encode: false });
    }
    const cookieList = Object.entries(def.cookies || {})
      .map(([k, v]) => k + '=' + v).join('; ');
    if (cookieList) headers['Cookie'] = cookieList;

    let payload = null;
    try {
      const res = await fetch(url, { headers, signal: AbortSignal.timeout(TIMEOUT_MS), redirect: 'manual' });
      if (res.status >= 300 && res.status < 400) {
        const location = res.headers && typeof res.headers.get === 'function' ? res.headers.get('location') : '';
        if (!location) {
          return { ok: false, results: [], error: { code: 'http', message: '「' + this.name + '」返回重定向但缺少 Location' } };
        }
        const redirected = new URL(location, url).toString();
        try {
          await ensureSafeEndpoint(redirected, def.allowPrivate === true);
        } catch (e) {
          return { ok: false, results: [], error: { code: 'bad-request', message: '「' + this.name + '」重定向地址未通过安全校验：' + e.message } };
        }
        return { ok: false, results: [], error: { code: 'http', message: '「' + this.name + '」返回重定向，请直接填写最终接口地址' } };
      }
      if (!res.ok) {
        return { ok: false, results: [], error: { code: 'http', message: '「' + this.name + '」响应异常（HTTP ' + res.status + '）' } };
      }
      try {
        payload = await res.json();
      } catch {
        return { ok: false, results: [], error: { code: 'parse', message: '「' + this.name + '」响应不是有效 JSON（请检查接口地址与返回格式）' } };
      }
    } catch (e) {
      if (e && e.name === 'TimeoutError') {
        return { ok: false, results: [], error: { code: 'timeout', message: '「' + this.name + '」请求超时' } };
      }
      return { ok: false, results: [], error: { code: 'network', message: '「' + this.name + '」请求失败：' + ((e && e.message) || String(e)) } };
    }

    const list = getPath(payload, def.resultsPath);
    if (!Array.isArray(list) || !list.length) {
      return { ok: true, results: [], error: { code: 'parse', message: '「' + this.name + '」未在 ' + (def.resultsPath || 'JSON 根部') + ' 找到结果数组' } };
    }

    const results = [];
    for (const item of list) {
      if (!item || typeof item !== 'object') continue;
      const imageUrl = String(getPath(item, def.fields.imageUrl) || '');
      if (!imageUrl) continue;
      const raw = {
        imageUrl,
        thumbnailUrl: def.fields.thumbnailUrl ? String(getPath(item, def.fields.thumbnailUrl) || '') : '',
        width: getPath(item, def.fields.width),
        height: getPath(item, def.fields.height),
        title: def.fields.title ? String(getPath(item, def.fields.title) || '') : '',
        sourceUrl: def.fields.sourceUrl ? String(getPath(item, def.fields.sourceUrl) || '') : '',
        author: def.fields.author ? String(getPath(item, def.fields.author) || '') : '',
        license: def.fields.license ? String(getPath(item, def.fields.license) || '') : '',
        mimeType: def.fields.mimeType ? String(getPath(item, def.fields.mimeType) || '') : '',
        source: def.fields.source ? String(getPath(item, def.fields.source) || '') : this.name
      };
      results.push(normalizeResult(raw, { provider: this.id, name: this.name, query }));
      if (results.length >= count) break;
    }
    if (!results.length) {
      return { ok: true, results: [], error: { code: 'parse', message: '「' + this.name + '」未解析到有效图片（请检查 fields 字段映射）' } };
    }
    return { ok: true, results, error: null };
  }
}

module.exports = {
  JsonApiProvider, validateSourceDef, sanitizeSourceDef, getPath, fillTemplate,
  _clearHostCache: () => hostCache.clear()
};
