// 图源管理路由（挂载于 /api/images/sources）：
//   自定义源 = 配置 image.sources[] 里的 JSON 定义（含默认启用源与用户添加的源）
//   内置源（代码实现）不可删除；预置模板只作「一键添加」的模板，不进配置
'use strict';
const express = require('express');
const { CONFIG_FILE, loadConfig, saveConfig } = require('../config.js');
const { validateSourceDef, sanitizeSourceDef, JsonApiProvider } = require('../providers/jsonApiProvider.js');
const { listBuiltins } = require('../providers/index.js');
const { presets } = require('../providers/presetSources.js');

function getSources() {
  const cfg = loadConfig();
  return Array.isArray(cfg.image.sources) ? cfg.image.sources : [];
}

function setSources(sources) {
  const cfg = loadConfig();
  cfg.image.sources = sources;
  saveConfig(CONFIG_FILE, cfg);
}

// 响应里 key 一律掩码（真实 Key 只在本机落盘加密存储）
function maskSource(s) {
  const { key, ...rest } = s || {};
  return { ...rest, key: key ? '****' : '' };
}

function sourcesRouter() {
  const router = express.Router();
  const builtins = () => listBuiltins().map((p) => ({ id: p.id, name: p.name, builtin: true }));

  // 全量视图：内置源 + 自定义源（含预置标记）+ 预置模板库
  router.get('/', (req, res) => {
    const presetIds = new Set(presets.map((p) => p.id));
    res.json({
      builtins: builtins(),
      custom: getSources().map((s) => ({ ...maskSource(s), builtin: false, preset: presetIds.has(s.id) })),
      presets: presets.map((p) => ({ ...p, key: '' }))
    });
  });

  // 新增 / 更新（按 id upsert）
  router.post('/', (req, res) => {
    const source = { ...((req.body || {}).source || {}) };
    const existing = getSources().find((s) => s.id === source.id);
    if (existing && source.key && String(source.key).includes('****') && existing.key) source.key = existing.key;
    const v = validateSourceDef(source);
    if (!v.ok) return res.status(400).json({ ok: false, errors: v.errors });
    if (builtins().some((b) => b.id === source.id)) {
      return res.status(400).json({ ok: false, errors: ['id 与内置源冲突：' + source.id] });
    }
    const clean = sanitizeSourceDef(source);
    const list = getSources();
    const idx = list.findIndex((s) => s.id === clean.id);
    // 掩码回写保护：更新已有源时若 key 是掩码，保留原 key
    if (idx >= 0 && clean.key && clean.key.includes('****')) {
      const prev = list[idx];
      if (prev && prev.key) clean.key = prev.key;
    }
    if (idx >= 0) list[idx] = clean; else list.push(clean);
    setSources(list);
    res.json({ ok: true, source: maskSource(clean) });
  });

  router.delete('/:id', (req, res) => {
    const id = req.params.id;
    if (builtins().some((b) => b.id === id)) {
      return res.status(400).json({ ok: false, errors: ['内置源不可删除'] });
    }
    setSources(getSources().filter((s) => s.id !== id));
    res.json({ ok: true });
  });

  // 批量导入：数组或单个对象；逐条校验，跳过无效/重复并报告
  router.post('/import', (req, res) => {
    const raw = (req.body || {}).sources;
    const arr = Array.isArray(raw) ? raw : (raw && typeof raw === 'object' ? [raw] : null);
    if (!arr) return res.status(400).json({ ok: false, errors: ['sources 必须是 JSON 数组或对象'] });
    const added = [];
    const skipped = [];
    const list = getSources();
    const builtinIds = new Set(builtins().map((b) => b.id));
    for (const src of arr) {
      const v = validateSourceDef(src);
      if (!v.ok) {
        skipped.push({ id: src && src.id, reason: v.errors.join('；') });
        continue;
      }
      if (builtinIds.has(src.id)) {
        skipped.push({ id: src.id, reason: '与内置源冲突' });
        continue;
      }
      const clean = sanitizeSourceDef(src);
      if (list.some((s) => s.id === clean.id)) {
        skipped.push({ id: clean.id, reason: '已存在（如需覆盖请编辑该源）' });
        continue;
      }
      list.push(clean);
      added.push({ id: clean.id });
    }
    if (added.length) setSources(list);
    res.json({ ok: true, added, skipped });
  });

  // 测试一个源定义（不保存）：用真实请求验证接口与字段映射
  router.post('/test', async (req, res) => {
    const { query } = req.body || {};
    const source = { ...((req.body || {}).source || {}) };
    const existing = getSources().find((s) => s.id === source.id);
    if (existing && source.key && String(source.key).includes('****') && existing.key) source.key = existing.key;
    const v = validateSourceDef(source);
    if (!v.ok) return res.status(400).json({ ok: false, errors: v.errors });
    const clean = sanitizeSourceDef(source);
    // 掩码 key 时用已保存的真实 key 测试
    if (clean.key && clean.key.includes('****')) {
      const prev = getSources().find((s) => s.id === clean.id);
      if (prev && prev.key) clean.key = prev.key;
    }
    const prov = new JsonApiProvider(clean);
    const r = await prov.search(query || '猫', { count: 3, page: 1 });
    const sample = r.results && r.results[0]
      ? { title: r.results[0].title, imageUrl: r.results[0].imageUrl, thumbnailUrl: r.results[0].thumbnailUrl }
      : null;
    res.json({ ok: r.ok, resultsCount: (r.results || []).length, error: r.error, sample });
  });

  return router;
}

module.exports = sourcesRouter;
