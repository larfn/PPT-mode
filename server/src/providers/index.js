// Provider 注册表：业务层通过 getProvider(id) 使用，不直接依赖任何 provider 的页面结构。
// 内置 provider（代码实现）静态注册；自定义源（配置里的 image.sources，JSON 定义）动态解析为 JsonApiProvider。
// 为未来 API Provider 预留：实现 ImageProvider 接口后在此 register 即可。
'use strict';
const { BaiduPageProvider } = require('./baiduPageProvider.js');
const { BingPageProvider } = require('./bingPageProvider.js');
const { QihooPageProvider } = require('./qihooPageProvider.js');
const { JsonApiProvider } = require('./jsonApiProvider.js');
const { loadConfig } = require('../config.js');

const providers = new Map();

function register(provider) {
  providers.set(provider.id, provider);
}

// 注册内置 provider（兼容实现：百度/必应/360 页面与 JSON 解析）
register(new BaiduPageProvider());
register(new BingPageProvider());
register(new QihooPageProvider());

// 内置（代码）源列表：不可被删除/覆盖
function listBuiltins() {
  return Array.from(providers.values()).map((p) => ({ id: p.id, name: p.name }));
}

// 配置里启用的自定义源定义（含预置模板被「添加」进配置后的条目）
function getEnabledCustomSources() {
  const cfg = loadConfig();
  const list = (cfg.image && Array.isArray(cfg.image.sources)) ? cfg.image.sources : [];
  return list.filter((s) => s && typeof s === 'object' && s.enabled !== false);
}

function getProvider(id) {
  if (providers.has(id)) return providers.get(id);
  const def = getEnabledCustomSources().find((s) => s.id === id);
  return def ? new JsonApiProvider(def) : null;
}

// 全部可用源（内置 + 启用的自定义源），供前端下拉/向导校验
function listProviders() {
  return [...listBuiltins(), ...getEnabledCustomSources().map((s) => ({ id: s.id, name: s.name }))];
}

module.exports = { getProvider, listProviders, listBuiltins, register };
