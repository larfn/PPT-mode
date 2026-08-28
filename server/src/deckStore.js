// 套版（Deck）存储：有序模板引用序列 + 每页填参规格
// 存储与模板平级：文档/PPT模板库/[分类]/套版名/deck.json + preview.jpg
// 套版引用模板（templateId/templateFolder/templateVersion），不内嵌模板快照；
// templateVersion 可选：固定版本号保证可复现（复用模板版本控制）。
'use strict';
const fs = require('node:fs');
const path = require('node:path');
const { getTemplate, templateRoot, safeId, safeFolder } = require('./templateStore.js');
const DECK_RECYCLE_FOLDER = '.套版回收站';
const DECK_RECYCLE_MARKER = '.deck-recycle.json';

function deckRoot() {
  return templateRoot();
}

// 解析套版目录：root / [folder] / [safeId]（与模板同根，防路径穿越）
function resolveDeckDir(id, folder) {
  const safe = safeId(id);
  if (!safe) return null;
  const f = safeFolder(folder);
  const r = path.resolve(deckRoot());
  const target = path.resolve(r, f, safe);
  if (target !== r && !target.startsWith(r + path.sep)) return null;
  return target;
}

function readDeckMeta(dir) {
  try { return JSON.parse(fs.readFileSync(path.join(dir, 'deck.json'), 'utf8')); }
  catch { return null; }
}

function readJsonFile(file) {
  try { return JSON.parse(fs.readFileSync(file, 'utf8')); } catch { return null; }
}

// 原子写（与模板一致）
function atomicWriteFile(file, content) {
  const tmp = file + '.tmp-' + process.pid + '-' + Date.now().toString(36);
  fs.writeFileSync(tmp, content, 'utf8');
  try { fs.renameSync(tmp, file); } catch (e) { try { fs.rmSync(tmp, { force: true }); } catch { /* ignore */ } throw e; }
}

function listDecks() {
  const dir = deckRoot();
  if (!fs.existsSync(dir)) return [];
  const out = [];
  const collect = (folderName, folderPath) => {
    for (const e of fs.readdirSync(folderPath, { withFileTypes: true })) {
      if (!e.isDirectory()) continue;
      const meta = readDeckMeta(path.join(folderPath, e.name));
      if (!meta) continue;
      out.push({
        id: e.name,
        folder: folderName,
        name: meta.name || e.name,
        pageCount: Array.isArray(meta.pages) ? meta.pages.length : 0,
        preview: '/api/decks/preview.png?folder=' + encodeURIComponent(folderName) + '&id=' + encodeURIComponent(e.name),
        updatedAt: meta.updatedAt || null
      });
    }
  };
  collect('', dir);
  for (const e of fs.readdirSync(dir, { withFileTypes: true })) {
    if (!e.isDirectory()) continue;
    if (e.name === DECK_RECYCLE_FOLDER || e.name === '.回收站') continue;
    if (readDeckMeta(path.join(dir, e.name))) continue;
    collect(e.name, path.join(dir, e.name));
  }
  return out.sort((a, b) => (b.updatedAt || '').localeCompare(a.updatedAt || ''));
}

// 保存/更新套版。校验：pages 引用的模板必须存在（templateVersion 固定时校验该版本）。
function saveDeck({ name, folder, deck, preview, assets = {} }) {
  const dir = resolveDeckDir(name, folder);
  if (!dir) throw new Error('invalid deck id');
  const id = safeId(name);
  const f = safeFolder(folder);
  const pages = Array.isArray(deck.pages) ? deck.pages : [];
  for (const p of pages) {
    if (!p || !p.templateId) throw new Error('页面缺少 templateId');
    const t = getTemplate(p.templateId, p.templateFolder || '');
    if (!t) throw new Error('页面引用的模板不存在：' + p.templateId);
    // 固定版本时校验版本存在
    if (p.templateVersion) {
      const { listVersions } = require('./templateStore.js');
      const lv = listVersions(p.templateId, p.templateFolder || '');
      const ok = lv && lv.versions.some((v) => v.versionId === String(p.templateVersion));
      if (!ok) throw new Error('页面引用的模板版本不存在：' + p.templateId + ' @ ' + p.templateVersion);
    }
  }
  fs.mkdirSync(dir, { recursive: true });
  const now = new Date().toISOString();
  const existing = readDeckMeta(dir);
  const meta = {
    ...deck,
    name, id, folder: f || undefined,
    schemaVersion: 1,
    createdAt: existing && existing.createdAt ? existing.createdAt : now,
    updatedAt: now
  };
  atomicWriteFile(path.join(dir, 'deck.json'), JSON.stringify(meta, null, 2));
  if (preview) {
    const m = preview.match(/^data:image\/(\w+);base64,(.*)$/);
    if (m) fs.writeFileSync(path.join(dir, 'preview.' + m[1]), Buffer.from(m[2], 'base64'));
  }
  return { id, name };
}

function getDeck(id, folder) {
  const dir = resolveDeckDir(id, folder);
  if (!dir) return null;
  const meta = readDeckMeta(dir);
  if (!meta) return null;
  return { name: meta.name, deck: meta };
}

function deleteDeck(id, folder) {
  const dir = resolveDeckDir(id, folder);
  if (!dir) return;
  if (!fs.existsSync(dir)) return;
  const f = safeFolder(folder);
  const rc = path.join(deckRoot(), DECK_RECYCLE_FOLDER, f);
  fs.mkdirSync(rc, { recursive: true });
  let target = path.join(rc, path.basename(dir));
  if (fs.existsSync(target)) target = path.join(rc, path.basename(dir) + '__' + Date.now().toString(36));
  const meta = readDeckMeta(dir) || {};
  atomicWriteFile(path.join(dir, DECK_RECYCLE_MARKER), JSON.stringify({
    id: safeId(id), folder: f, name: meta.name || safeId(id), deletedAt: new Date().toISOString()
  }, null, 2));
  fs.renameSync(dir, target);
}

function deckRecycleDir() {
  return path.join(deckRoot(), DECK_RECYCLE_FOLDER);
}

function resolveDeckRecycleEntry(entryId) {
  if (!entryId || typeof entryId !== 'string' || !entryId.trim()) return null;
  const cleaned = String(entryId).replace(/^[\/\\]+/, '').trim();
  if (!cleaned || cleaned === '.' || cleaned === '..') return null;
  const rc = path.resolve(deckRecycleDir());
  const target = path.resolve(rc, cleaned);
  if (target === rc || !target.startsWith(rc + path.sep)) return null;
  return target;
}

function previewUrlForRecycle(entryId) {
  return '/api/decks/recycle/preview.png?entryId=' + encodeURIComponent(entryId);
}

function listDeckRecycleBin() {
  const rc = deckRecycleDir();
  if (!fs.existsSync(rc)) return [];
  const out = [];
  const collect = (folderName, folderPath) => {
    for (const e of fs.readdirSync(folderPath, { withFileTypes: true })) {
      if (!e.isDirectory()) continue;
      const dir = path.join(folderPath, e.name);
      if (!fs.existsSync(path.join(dir, DECK_RECYCLE_MARKER)) && !readDeckMeta(dir)) continue;
      const marker = readJsonFile(path.join(dir, DECK_RECYCLE_MARKER)) || {};
      const meta = readDeckMeta(dir) || {};
      const entryId = (folderName ? folderName + '/' : '') + e.name;
      out.push({
        entryId,
        id: e.name,
        folder: marker.folder !== undefined ? marker.folder : (folderName || ''),
        name: marker.name || meta.name || e.name,
        deletedAt: marker.deletedAt || null,
        pageCount: Array.isArray(meta.pages) ? meta.pages.length : 0,
        preview: previewUrlForRecycle(entryId)
      });
    }
  };
  collect('', rc);
  for (const e of fs.readdirSync(rc, { withFileTypes: true })) {
    if (!e.isDirectory()) continue;
    if (readDeckMeta(path.join(rc, e.name))) continue;
    collect(e.name, path.join(rc, e.name));
  }
  return out.sort((a, b) => (b.deletedAt || '').localeCompare(a.deletedAt || ''));
}

function restoreDeck(entryId) {
  const src = resolveDeckRecycleEntry(entryId);
  if (!src || !fs.existsSync(src)) throw new Error('回收站条目不存在');
  const marker = readJsonFile(path.join(src, DECK_RECYCLE_MARKER)) || {};
  const origFolder = marker.folder !== undefined ? marker.folder : '';
  const id = safeId(marker.id || path.basename(src));
  const root = path.resolve(deckRoot());
  const dest = path.resolve(root, origFolder ? safeFolder(origFolder) : '', id);
  if (dest !== root && !dest.startsWith(root + path.sep)) throw new Error('非法恢复目标');
  if (fs.existsSync(dest)) throw new Error('目标位置已存在同名套版');
  fs.mkdirSync(path.dirname(dest), { recursive: true });
  fs.renameSync(src, dest);
  try { fs.rmSync(path.join(dest, DECK_RECYCLE_MARKER), { force: true }); } catch { /* ignore */ }
  return { ok: true, id, folder: origFolder };
}

function purgeDeck(entryId) {
  const target = resolveDeckRecycleEntry(entryId);
  if (!target || target === deckRecycleDir()) return { ok: false, error: '非法条目' };
  if (!fs.existsSync(target)) return { ok: false, error: '回收站条目不存在' };
  fs.rmSync(target, { recursive: true, force: true });
  return { ok: true };
}

function emptyDeckRecycleBin() {
  const rc = deckRecycleDir();
  if (!fs.existsSync(rc)) return { ok: true, removed: 0 };
  let removed = 0;
  for (const e of fs.readdirSync(rc, { withFileTypes: true })) {
    const p = path.join(rc, e.name);
    try {
      if (e.isDirectory()) fs.rmSync(p, { recursive: true, force: true });
      else fs.rmSync(p, { force: true });
      removed++;
    } catch { /* ignore */ }
  }
  return { ok: true, removed };
}

module.exports = {
  listDecks, saveDeck, getDeck, deleteDeck, deckRoot, resolveDeckDir,
  listDeckRecycleBin, restoreDeck, purgeDeck, emptyDeckRecycleBin, deckRecycleDir, resolveDeckRecycleEntry
};
