const fs = require('node:fs');
const path = require('node:path');
const { protectSecret, unprotectSecret } = require('./security.js');
const { DATA_DIR, CONFIG_FILE } = require('./paths.js');
const { defaultEnabledSources } = require('./providers/presetSources.js');

// 敏感字段：保存时用 Windows DPAPI 加密（dpapi:<base64>），读取时透明解密。
// 内存缓存（按文件 mtime）避免每次读取都 spawn PowerShell。
// 除标量字段（apiKey/bingApiKey）外，自定义图源里的 key 同样加密保存。
const SECRET_KEYS = ['apiKey', 'bingApiKey'];

function defaultConfig() {
  return {
    text: { baseUrl: 'https://api.deepseek.com', apiKey: '', model: 'deepseek-chat' },
    image: {
      provider: 'baidu_page',
      bingApiKey: '',
      pageSize: 9,
      sources: JSON.parse(JSON.stringify(defaultEnabledSources)) // 默认初始搜图源（可编辑/删除）
    },
    analyze: { enabled: false }, // AI 模板分析总开关（配置页勾选才启用）
    highlight: { color: '#FF0000', durationMs: 1500 },
    ui: { fontSize: 14, language: 'zh' }
  };
}

function mergeConfig(parsed) {
  const def = defaultConfig();
  return {
    text: { ...def.text, ...(parsed.text || {}) },
    image: { ...def.image, ...(parsed.image || {}) },
    analyze: { ...def.analyze, ...(parsed.analyze || {}) },
    highlight: { ...def.highlight, ...(parsed.highlight || {}) },
    ui: { ...def.ui, ...(parsed.ui || {}) }
  };
}

// 自定义图源的 key 字段加密/解密（幂等：已是 dpapi: 或空值保持不变）
function transformSourceKeys(sources, fn) {
  if (!Array.isArray(sources)) return sources;
  for (const s of sources) {
    if (s && typeof s === 'object' && typeof s.key === 'string' && s.key) s.key = fn(s.key);
  }
  return sources;
}

// 深拷贝 + 加密敏感字段（已是 dpapi: 或空值保持不变，幂等）
function encryptSecrets(cfg) {
  const out = JSON.parse(JSON.stringify(cfg));
  for (const section of ['text', 'image']) {
    const s = out[section];
    if (!s || typeof s !== 'object') continue;
    for (const key of SECRET_KEYS) {
      if (typeof s[key] === 'string' && s[key]) s[key] = protectSecret(s[key]);
    }
    if (section === 'image') transformSourceKeys(s.sources, protectSecret);
  }
  return out;
}

// 解密敏感字段（内存中始终是明文，调用方行为与旧版完全一致）
function decryptSecrets(cfg) {
  for (const section of ['text', 'image']) {
    const s = cfg[section];
    if (!s || typeof s !== 'object') continue;
    for (const key of SECRET_KEYS) {
      if (typeof s[key] === 'string') s[key] = unprotectSecret(s[key]);
    }
    if (section === 'image') transformSourceKeys(s.sources, unprotectSecret);
  }
  return cfg;
}

// 掩码回写保护：前端提交的自定义源里 key 为掩码（****）时，保留原 key 不回写
function mergeSourcesMasked(currentSources, incomingSources) {
  if (!Array.isArray(incomingSources)) return currentSources;
  const cur = Array.isArray(currentSources) ? currentSources : [];
  return incomingSources.map((s) => {
    const prev = cur.find((c) => c && c.id === s.id);
    if (prev && prev.key && typeof s.key === 'string' && s.key.includes('****')) {
      s.key = prev.key;
    }
    return s;
  });
}

// 内存缓存：path -> { mtimeMs, cfg }（cfg 为已解密明文）。mtime 变化即失效，
// 保证外部修改（手动编辑/其他进程）能被重新读取。
const cfgCache = new Map();

function loadConfig(filePath = CONFIG_FILE) {
  try {
    const st = fs.statSync(filePath);
    const hit = cfgCache.get(filePath);
    if (hit && hit.mtimeMs === st.mtimeMs) return hit.cfg;
    const raw = fs.readFileSync(filePath, 'utf8');
    const parsed = JSON.parse(raw);
    const cfg = decryptSecrets(mergeConfig(parsed));
    cfgCache.set(filePath, { mtimeMs: st.mtimeMs, cfg });
    return cfg;
  } catch {
    return defaultConfig();
  }
}

function saveConfig(filePath = CONFIG_FILE, cfg) {
  fs.mkdirSync(path.dirname(filePath), { recursive: true });
  const encrypted = encryptSecrets(cfg);
  fs.writeFileSync(filePath, JSON.stringify(encrypted, null, 2), 'utf8');
  // 更新缓存为明文 cfg，后续 loadConfig 免解密
  try {
    const st = fs.statSync(filePath);
    cfgCache.set(filePath, { mtimeMs: st.mtimeMs, cfg: JSON.parse(JSON.stringify(cfg)) });
  } catch { /* 缓存失败不影响 */ }
}

function maskKey(key) {
  if (!key) return '';
  if (key.length <= 6) return '****';
  return key.slice(0, 3) + '****';
}

module.exports = { DATA_DIR, CONFIG_FILE, defaultConfig, loadConfig, saveConfig, maskKey, mergeSourcesMasked };
