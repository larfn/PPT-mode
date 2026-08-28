// 模板存储（含模板版本控制 Template Versioning）
//
// 文件结构（保持旧模板目录兼容）：
//   模板目录/
//     template.json        ← 当前版本镜像（所有旧读取逻辑不变）+ 版本元数据（version/versionId/createdAt/updatedAt/changeNote）
//     preview.jpg          ← 当前预览（兼容旧逻辑）
//     versions/
//       v1.json / v1.jpg   ← 版本 1 的模板数据与预览（每版本独立）
//       v2.json / v2.jpg   ← 版本 2 …
//
// 核心约定：
//   - 版本号从 1 递增；versionId = 'v' + 版本号（v1/v2/…）
//   - template.json 永远是「当前版本」的镜像 → 旧代码、MCP、向导全部零改动可用
//   - 每次保存默认创建新版本（v1 → v2），绝不静默覆盖；updateCurrent=true 时修正当前版本（前端后台补存样式/预览用）
//   - 旧模板（无 versions/）读取时惰性自动迁移为 v1/current，无需用户操作；迁移失败不影响读取
//   - 所有写操作用「临时文件 + rename」原子写入；进程内写锁串行化，防并发版本号冲突
//   - template.json 损坏时自动从 versions/ 最新版本恢复
'use strict';
const fs = require('node:fs');
const os = require('node:os');
const path = require('node:path');
const { normalizeTemplate } = require('./semantic.js');

let root = null;

function templateRoot() {
  if (root) return root;
  return process.env.PPT_TEMPLATE_ROOT || path.join(os.homedir(), 'Documents', 'PPT模板库');
}

function setTemplateRoot(dir) { root = dir; }

// ---------- 原子写入：临时文件 → rename（防写入中断/JSON 损坏/异常关机） ----------
function atomicWriteFile(file, content) {
  const tmp = file + '.tmp-' + process.pid + '-' + Date.now().toString(36) + '-' + Math.random().toString(36).slice(2, 6);
  fs.writeFileSync(tmp, content, 'utf8');
  try {
    fs.renameSync(tmp, file); // Windows 下 Node rename 可覆盖已存在文件（MoveFileEx 语义）
  } catch (e) {
    try { fs.rmSync(tmp, { force: true }); } catch { /* ignore */ }
    throw e;
  }
}
function atomicWriteJson(file, obj) { atomicWriteFile(file, JSON.stringify(obj, null, 2)); }
function atomicCopyFile(src, dst) {
  const tmp = dst + '.tmp-' + process.pid + '-' + Date.now().toString(36) + '-' + Math.random().toString(36).slice(2, 6);
  fs.copyFileSync(src, tmp);
  try { fs.renameSync(tmp, dst); } catch (e) { try { fs.rmSync(tmp, { force: true }); } catch { /* ignore */ } throw e; }
}

// 并发说明：本模块全部为同步 fs 操作，Node 单线程事件循环保证
// 「读取当前版本号 + 写入新版本」在一个同步块内完成，单进程内不会并发交错；
// 跨进程（两个后端实例）由 createVersionFileExclusive 的 O_EXCL 占位防双写。

// ---------- 路径与安全 ----------
function safeId(name) {
  const safe = name.trim().replace(/[\\/:*?"<>|]/g, '_').slice(0, 60);
  return safe === '.' || safe === '..' ? '' : safe;
}
function safeFolder(folder) {
  const f = (folder || '').trim().replace(/[\\/:*?"<>|]/g, '_').slice(0, 60);
  return f === '.' || f === '..' ? '' : f;
}
function resolveDir(id, folder) {
  const safe = safeId(id);
  if (!safe) return null;
  const f = safeFolder(folder);
  const r = path.resolve(templateRoot());
  const target = path.resolve(r, f, safe);
  if (target !== r && !target.startsWith(r + path.sep)) return null;
  return target;
}
function versionsDir(dir) { return path.join(dir, 'versions'); }
const VERSION_ID_RE = /^v[1-9]\d*$/;

// ---------- 读取 ----------
function readMeta(dir) {
  try { return JSON.parse(fs.readFileSync(path.join(dir, 'template.json'), 'utf8')); }
  catch { return null; }
}
function parseVersionFile(dir, versionId) {
  if (!VERSION_ID_RE.test(versionId || '')) return null;
  const f = path.join(versionsDir(dir), versionId + '.json');
  if (!fs.existsSync(f)) return null;
  try { return JSON.parse(fs.readFileSync(f, 'utf8')); } catch { return null; }
}
function listVersionIds(dir) {
  const vd = versionsDir(dir);
  if (!fs.existsSync(vd)) return [];
  return fs.readdirSync(vd)
    .filter((f) => VERSION_ID_RE.test(f.replace(/\.json$/, '')) && f.endsWith('.json'))
    .map((f) => f.replace(/\.json$/, ''))
    .sort((a, b) => Number(a.slice(1)) - Number(b.slice(1)));
}
function findPreviewFile(dir, versionId) {
  const base = versionId ? versionsDir(dir) : dir;
  const names = versionId ? [versionId + '.png', versionId + '.jpg', versionId + '.jpeg'] : ['preview.png', 'preview.jpg', 'preview.jpeg'];
  for (const n of names) {
    const f = path.join(base, n);
    if (fs.existsSync(f)) return f;
  }
  return null;
}
function removeCurrentPreview(dir) {
  for (const n of ['preview.png', 'preview.jpg', 'preview.jpeg']) {
    try { fs.rmSync(path.join(dir, n), { force: true }); } catch { /* ignore */ }
  }
}

// ---------- 旧模板惰性迁移（幂等；失败不影响读取；迁移在写锁内执行） ----------
// template.json 无版本结构 → 生成 versions/v1.json 快照 + preview 副本 + 写回版本元数据
function migrateLegacyLocked(dir) {
  const meta = readMeta(dir);
  if (!meta) return null;
  if (meta.versionId && VERSION_ID_RE.test(meta.versionId)) {
    // 已是新结构：绝不重新迁移（即使 versions/ 缺文件，可能是用户删除了该版本）。
    // 「当前版本」文件异常缺失（手动删除/损坏）时，从当前镜像自动补写，保证当前版本始终可恢复。
    const curFile = path.join(versionsDir(dir), meta.versionId + '.json');
    if (!fs.existsSync(curFile)) {
      try {
        fs.mkdirSync(versionsDir(dir), { recursive: true });
        createVersionFileExclusive(curFile, {
          ...meta, version: meta.version, versionId: meta.versionId,
          updatedAt: meta.updatedAt || new Date().toISOString()
        });
      } catch { /* 补写失败不影响读取 */ }
    }
    return meta;
  }
  // 旧模板（无版本元数据）→ 惰性迁移为 v1/current
  const vd = versionsDir(dir);
  const v1File = path.join(vd, 'v1.json');
  try {
    fs.mkdirSync(vd, { recursive: true });
    const now = new Date().toISOString();
    const createdAt = meta.createdAt || now;
    const v1 = { ...meta, version: 1, versionId: 'v1', createdAt, updatedAt: now, changeNote: meta.changeNote || '首次迁移（旧模板）' };
    if (!fs.existsSync(v1File)) {
      try { createVersionFileExclusive(v1File, v1); } catch { /* 并发迁移：已有则跳过 */ }
    }
    const previewFile = findPreviewFile(dir);
    if (previewFile) { try { atomicCopyFile(previewFile, path.join(vd, 'v1' + path.extname(previewFile))); } catch { /* 预览副本失败不影响 */ } }
    atomicWriteJson(path.join(dir, 'template.json'), { ...v1, id: meta.id, folder: meta.folder });
    return v1;
  } catch {
    return meta; // 迁移失败：按旧模板继续使用，下次再试
  }
}

// ---------- 跨进程防双写：O_EXCL 占位创建版本文件，再原子写入内容 ----------
function createVersionFileExclusive(file, obj) {
  let fd;
  try { fd = fs.openSync(file, 'wx'); } catch (e) { if (e.code === 'EEXIST') return false; throw e; }
  try { fs.closeSync(fd); } catch { /* ignore */ }
  atomicWriteFile(file, JSON.stringify(obj, null, 2));
  return true;
}

// ---------- 预览写入（当前 + 版本快照，均为原子写；失败不影响模板保存） ----------
function writePreviewFiles(dir, preview, versionId) {
  if (!preview) return;
  const m = preview.match(/^data:image\/(\w+);base64,(.*)$/);
  if (!m) return;
  const bytes = Buffer.from(m[2], 'base64');
  try {
    if (versionId) {
      const vd = versionsDir(dir);
      fs.mkdirSync(vd, { recursive: true });
      atomicWriteFile(path.join(vd, versionId + '.' + m[1]), bytes);
    }
    atomicWriteFile(path.join(dir, 'preview.' + m[1]), bytes);
  } catch { /* 预览写入失败不影响模板本体 */ }
}

// ---------- 保存 ----------
// 默认（updateCurrent=false）：创建新版本（v1 → v2 → v3…），绝不静默覆盖
// updateCurrent=true：修正当前版本内容（前端「保存后后台补存精确样式/预览图」场景，不产生新版本）
function saveTemplate({ name, folder, template, preview, assets = {}, changeNote, updateCurrent = false }) {
  const dir = resolveDir(name, folder);
  if (!dir) throw new Error('invalid template id');
  const id = safeId(name);
  const f = safeFolder(folder);
  fs.mkdirSync(path.join(dir, 'assets'), { recursive: true });
  {
    const now = new Date().toISOString();
    const clean = normalizeTemplate(template);
    const existing = migrateLegacyLocked(dir); // 旧模板自动迁移；新模板为 null
    const createdAt = (existing && existing.createdAt) || now;
    const note = typeof changeNote === 'string' && changeNote.trim() ? changeNote.trim() : undefined;

    if (updateCurrent) {
      const cur = existing;
      if (!cur || !VERSION_ID_RE.test(cur.versionId || '')) throw new Error('updateCurrent 需要已存在的当前版本');
      const vId = cur.versionId;
      const vdata = { ...clean, name, id, folder: f || undefined, version: cur.version, versionId: vId, createdAt: cur.createdAt || createdAt, updatedAt: now, changeNote: note !== undefined ? note : cur.changeNote };
      atomicWriteFile(path.join(versionsDir(dir), vId + '.json'), JSON.stringify(vdata, null, 2));
      writePreviewFiles(dir, preview, vId);
      atomicWriteJson(path.join(dir, 'template.json'), vdata);
      return { id, name, version: cur.version, versionId: vId, created: false };
    }

    // 新建版本：版本号 = 当前 + 1；wx 占位防跨进程冲突，冲突则递增重试
    let v = (existing && Number.isInteger(existing.version) && existing.version > 0 ? existing.version : 0) + 1;
    let vId = 'v' + v;
    const vd = versionsDir(dir);
    fs.mkdirSync(vd, { recursive: true });
    const base = { ...clean, name, id, folder: f || undefined, createdAt, updatedAt: now, changeNote: note };
    for (let i = 0; i < 20; i++) {
      if (createVersionFileExclusive(path.join(vd, vId + '.json'), { ...base, version: v, versionId: vId })) break;
      v += 1; vId = 'v' + v;
    }
    writePreviewFiles(dir, preview, vId);
    atomicWriteJson(path.join(dir, 'template.json'), { ...base, version: v, versionId: vId });
    invalidateTemplateCache();
    return { id, name, version: v, versionId: vId, created: true };
  }
}

// ---------- 版本列表 ----------
function listVersions(id, folder) {
  const dir = resolveDir(id, folder);
  if (!dir) return null;
  {
    const meta = migrateLegacyLocked(dir);
    if (!meta) return null;
    const cur = meta.versionId;
    const versions = listVersionIds(dir).map((vId) => {
      const v = parseVersionFile(dir, vId) || {};
      return {
        version: Number(vId.slice(1)),
        versionId: vId,
        createdAt: v.createdAt || null,
        updatedAt: v.updatedAt || null,
        changeNote: v.changeNote || null,
        isCurrent: vId === cur
      };
    });
    return { currentVersion: meta.version, currentVersionId: cur, versions };
  }
}

// ---------- 读取单个版本 ----------
function getVersion(id, folder, versionId) {
  const dir = resolveDir(id, folder);
  if (!dir) return null;
  {
    const meta = migrateLegacyLocked(dir);
    if (!meta) return null;
    const v = parseVersionFile(dir, versionId);
    if (!v) return null;
    const previewFile = findPreviewFile(dir, versionId);
    return {
      version: normalizeTemplate(v),
      previewUrl: previewFile
        ? '/api/templates/preview.png?folder=' + encodeURIComponent(safeFolder(folder)) + '&id=' + encodeURIComponent(id) + '&version=' + encodeURIComponent(versionId)
        : null,
      isCurrent: meta.versionId === versionId
    };
  }
}

// ---------- 恢复 / 设为当前版本（同一实现：版本内容成为当前镜像，preview 同步，保证不错配） ----------
function restoreVersion(id, folder, versionId) {
  const dir = resolveDir(id, folder);
  if (!dir) throw new Error('invalid template id');
  {
    const meta = migrateLegacyLocked(dir);
    if (!meta) throw new Error('模板不存在');
    if (!VERSION_ID_RE.test(versionId || '')) throw new Error('非法版本号：' + versionId);
    const v = parseVersionFile(dir, versionId);
    if (!v) throw new Error('版本不存在：' + versionId);
    const now = new Date().toISOString();
    const ver = Number(versionId.slice(1));
    atomicWriteJson(path.join(dir, 'template.json'), {
      ...normalizeTemplate(v), name: meta.name, id: meta.id, folder: meta.folder,
      version: ver, versionId, createdAt: meta.createdAt || v.createdAt || now, updatedAt: now,
      changeNote: typeof v.changeNote === 'string' ? v.changeNote : undefined
    });
    const vp = findPreviewFile(dir, versionId);
    if (vp) {
      try { atomicCopyFile(vp, path.join(dir, 'preview' + path.extname(vp))); } catch { /* 预览复制失败：删除当前预览避免错配 */ removeCurrentPreview(dir); }
    } else {
      removeCurrentPreview(dir); // 该版本无预览：删掉当前预览，宁可无图也不错配
    }
    invalidateTemplateCache();
    return { ok: true, version: ver, versionId };
  }
}
function setCurrentVersion(id, folder, versionId) {
  // 「设为当前版本」与「恢复」在当前镜像模型下等价：template.json 始终是当前版本内容，
  // 只改标记会造成内容与标记不一致，因此统一为把版本内容设为当前。
  return restoreVersion(id, folder, versionId);
}

// ---------- 删除单个版本（与删除整个模板严格区分） ----------
function deleteVersion(id, folder, versionId) {
  const dir = resolveDir(id, folder);
  if (!dir) throw new Error('invalid template id');
  {
    const meta = migrateLegacyLocked(dir);
    if (!meta) throw new Error('模板不存在');
    if (!VERSION_ID_RE.test(versionId || '')) throw new Error('非法版本号：' + versionId);
    const ids = listVersionIds(dir);
    if (!ids.includes(versionId)) throw new Error('版本不存在：' + versionId);
    // 先查唯一版本：唯一版本无论是否当前都不能删（没有可切换目标，应删整个模板）
    if (ids.length <= 1) throw new Error('模板至少保留一个版本；删除整个模板请使用「删除模板」');
    if (meta.versionId === versionId) throw new Error('不能删除当前版本，请先「设为当前版本」切换到其他版本');
    fs.rmSync(path.join(versionsDir(dir), versionId + '.json'), { force: true });
    for (const n of [versionId + '.png', versionId + '.jpg', versionId + '.jpeg']) {
      try { fs.rmSync(path.join(versionsDir(dir), n), { force: true }); } catch { /* ignore */ }
    }
    invalidateTemplateCache();
    return { ok: true };
  }
}

// ---------- 模板读取（当前版本）：损坏自动恢复 ----------
// 定位模板目录：先按指定 folder 找；folder 未指定（或为空）时，全局扫描模板库（含子文件夹）
// 按 id（目录名）兜底查找 —— 解决 AI/MCP 场景下用户没传 folder 时模板明明在子文件夹却 not found 的问题。
// 注意：显式指定了 folder 且该位置没有时不做跨文件夹猜测（模板 id 可能在多个文件夹重名，避免歧义）。
function findTemplateDir(id, folder) {
  // 注意：只用「目录存在」判定（不要求 template.json 可读）——损坏的模板要交给 getTemplate 的
  // recoverFromVersions 从版本文件恢复；若这里因 readMeta 失败就返回 null，恢复逻辑永远走不到。
  const direct = resolveDir(id, folder);
  if (direct && fs.existsSync(direct)) return direct;
  if (folder && String(folder).trim()) return null;
  const safe = safeId(id);
  if (!safe) return null;
  const r = path.resolve(templateRoot());
  if (!fs.existsSync(r)) return null;
  const found = [];
  const scan = (dir, depth) => {
    if (depth > 5 || found.length) return;
    let entries;
    try { entries = fs.readdirSync(dir, { withFileTypes: true }); } catch { return; }
    for (const e of entries) {
      if (!e.isDirectory()) continue;
      if (e.name === RECYCLE_NAME) continue; // 跳过回收站（已删除的模板不该被全局查找命中）
      const p = path.join(dir, e.name);
      if (e.name === safe) { found.push(p); return; }
      scan(p, depth + 1);
    }
  };
  scan(r, 0);
  return found[0] || null;
}
function getTemplate(id, folder) {
  const dir = findTemplateDir(id, folder);
  if (!dir) return null;
  {
    let meta = readMeta(dir);
    if (!meta) {
      // template.json 损坏/缺失 → 尝试从版本文件恢复（最新版本优先）
      meta = recoverFromVersions(dir);
      if (!meta) return null;
    }
    meta = migrateLegacyLocked(dir) || meta; // 旧模板自动迁移（幂等）
    return { name: meta.name, template: normalizeTemplate(meta) };
  }
}

function recoverFromVersions(dir) {
  const ids = listVersionIds(dir);
  for (let i = ids.length - 1; i >= 0; i--) {
    const v = parseVersionFile(dir, ids[i]);
    if (!v) continue;
    try {
      const vp = findPreviewFile(dir, ids[i]);
      if (vp) { try { atomicCopyFile(vp, path.join(dir, 'preview' + path.extname(vp))); } catch { /* ignore */ } }
      const meta = { ...v, updatedAt: new Date().toISOString() };
      atomicWriteJson(path.join(dir, 'template.json'), meta);
      return meta;
    } catch {
      return v; // 重建失败也返回数据，保证只读路径可用
    }
  }
  return null;
}

// ---------- 列表 ----------
// 模板库列表缓存（性能）：以「根目录条目 mtime + 子模板目录 mtime」为 key，
// 内容没变化时直接返回上次结果，避免每次进库都全量读每个 template.json（模板多时显著提速）。
// 写操作（保存/删除/恢复版本等）会改 mtime → key 自动失效；关键写操作也显式调用 invalidateTemplateCache 双保险。
let templateListCache = null;
let templateListCacheKey = null;
function invalidateTemplateCache() { templateListCache = null; templateListCacheKey = null; }

function templateListKey() {
  const dir = templateRoot();
  const parts = [];
  const push = (n, mtimeMs) => parts.push(n + ':' + Math.round(mtimeMs));
  for (const e of fs.readdirSync(dir, { withFileTypes: true })) {
    if (e.name === RECYCLE_NAME) continue;
    const p = path.join(dir, e.name);
    let st;
    try { st = fs.statSync(p); } catch { continue; }
    push(e.name, st.mtimeMs);
    if (e.isDirectory()) {
      // 子模板目录 mtime：template.json/preview 更新会反映到目录 mtime
      for (const sub of fs.readdirSync(p, { withFileTypes: true })) {
        if (sub.isDirectory()) {
          try { push(sub.name, fs.statSync(path.join(p, sub.name)).mtimeMs); } catch { /* ignore */ }
        }
      }
    }
  }
  return parts.join('|');
}

function computeListTemplates() {
  const dir = templateRoot();
  if (!fs.existsSync(dir)) return [];
  const out = [];
  const collect = (folderName, folderPath) => {
    for (const e of fs.readdirSync(folderPath, { withFileTypes: true })) {
      if (!e.isDirectory()) continue;
      if (e.name === RECYCLE_NAME) continue; // 回收站不列入模板库列表
      const meta = readMeta(path.join(folderPath, e.name));
      if (!meta) continue; // 非模板目录 / 损坏模板跳过（getTemplate 时自动恢复）
      out.push({
        id: e.name,
        folder: folderName,
        name: meta?.name || e.name,
        version: Number.isInteger(meta?.version) && meta.version > 0 ? meta.version : undefined,
        preview: '/api/templates/preview.png?folder=' + encodeURIComponent(folderName) + '&id=' + encodeURIComponent(e.name),
        updatedAt: meta?.updatedAt || null,
        // 前端据此判断是否需要在模板库自动生成结构示意图预览（缺失时为 false）
        hasPreview: ['preview.png', 'preview.jpg', 'preview.jpeg'].some((n) => fs.existsSync(path.join(folderPath, e.name, n)))
      });
    }
  };
  collect('', dir);
  for (const e of fs.readdirSync(dir, { withFileTypes: true })) {
    if (!e.isDirectory()) continue;
    if (e.name === RECYCLE_NAME) continue;
    if (readMeta(path.join(dir, e.name))) continue;
    collect(e.name, path.join(dir, e.name));
  }
  return out.sort((a, b) => (b.updatedAt || '').localeCompare(a.updatedAt || ''));
}

function listTemplates() {
  const dir = templateRoot();
  if (!fs.existsSync(dir)) return [];
  const key = templateListKey();
  if (templateListCache && templateListCacheKey === key) return templateListCache;
  const result = computeListTemplates();
  templateListCache = result;
  templateListCacheKey = key;
  return result;
}

function listFolders() {
  const dir = templateRoot();
  if (!fs.existsSync(dir)) return [];
  const out = [];
  for (const e of fs.readdirSync(dir, { withFileTypes: true })) {
    if (!e.isDirectory()) continue;
    if (e.name === RECYCLE_NAME) continue; // 回收站不显示为分类文件夹
    const folderPath = path.join(dir, e.name);
    let count = 0;
    for (const sub of fs.readdirSync(folderPath, { withFileTypes: true })) {
      if (sub.isDirectory() && readMeta(path.join(folderPath, sub.name))) count++;
    }
    if (count) out.push({ name: e.name, count });
  }
  return out.sort((a, b) => a.name.localeCompare(b.name, 'zh'));
}

// ---------- 回收站（P2-F）：删除模板 = 移入 .回收站（可恢复；版本控制已兜底，此为防误删补强） ----------
// 文件结构：文档/PPT模板库/.回收站/<原folder>/<模板id>/（保留原目录全部内容：template.json + versions/ + preview）
// 目录内写 .recycle.json 标记（原始 id/folder/name/deletedAt），恢复时按标记回原位置。
const RECYCLE_NAME = '.回收站';
const RECYCLE_MARKER = '.recycle.json';

function recycleDir() { return path.join(templateRoot(), RECYCLE_NAME); }
function readJsonFile(file) {
  try { return JSON.parse(fs.readFileSync(file, 'utf8')); } catch { return null; }
}
// 安全解析回收站条目路径（entryId 形如 'folder/id' 或 'id'；防穿越）
function resolveRecycleEntry(entryId) {
  if (!entryId || typeof entryId !== 'string' || !entryId.trim()) return null;
  const cleaned = String(entryId).replace(/^[\/\\]+/, '').trim();
  if (!cleaned || cleaned === '.' || cleaned === '..') return null;
  const rc = path.resolve(recycleDir());
  const target = path.resolve(rc, cleaned);
  if (target === rc || !target.startsWith(rc + path.sep)) return null;
  return target;
}
// 删除模板 → 移动到 .回收站（同卷 rename 原子移动；目标同名时加时间戳后缀）
function deleteTemplate(id, folder) {
  const dir = resolveDir(id, folder);
  if (!dir || !fs.existsSync(dir)) return;
  const rc = recycleDir();
  const f = safeFolder(folder);
  fs.mkdirSync(path.join(rc, f), { recursive: true });
  let target = path.join(rc, f, path.basename(dir));
  if (fs.existsSync(target)) target = path.join(rc, f, path.basename(dir) + '__' + Date.now().toString(36));
  try {
    const meta = readMeta(dir) || {};
    atomicWriteJson(path.join(dir, RECYCLE_MARKER), {
      id: safeId(id), folder: f, name: meta.name || safeId(id), deletedAt: new Date().toISOString()
    });
  } catch { /* marker 写失败不影响移动 */ }
  fs.renameSync(dir, target);
  invalidateTemplateCache();
}

// 列出回收站条目（含原位置、删除时间、预览）
function listRecycleBin() {
  const rc = recycleDir();
  if (!fs.existsSync(rc)) return [];
  const out = [];
  const collect = (folderName, folderPath) => {
    for (const e of fs.readdirSync(folderPath, { withFileTypes: true })) {
      if (!e.isDirectory()) continue;
      const dir = path.join(folderPath, e.name);
      // 只有模板目录（有 .recycle.json 标记或 template.json）才算条目；分类文件夹等跳过
      if (!fs.existsSync(path.join(dir, RECYCLE_MARKER)) && !readMeta(dir)) continue;
      const marker = readJsonFile(path.join(dir, RECYCLE_MARKER)) || {};
      const meta = readMeta(dir) || {};
      out.push({
        entryId: (folderName ? folderName + '/' : '') + e.name,
        id: e.name,
        folder: marker.folder !== undefined ? marker.folder : (folderName || ''),
        name: marker.name || meta.name || e.name,
        deletedAt: marker.deletedAt || null,
        preview: '/api/templates/preview.png?folder=' + encodeURIComponent(RECYCLE_NAME) + '&id=' + encodeURIComponent(e.name)
      });
    }
  };
  collect('', rc);
  for (const e of fs.readdirSync(rc, { withFileTypes: true })) {
    if (!e.isDirectory()) continue;
    if (readMeta(path.join(rc, e.name))) continue; // 根级模板目录（collect('') 已处理）
    collect(e.name, path.join(rc, e.name));
  }
  return out.sort((a, b) => (b.deletedAt || '').localeCompare(a.deletedAt || ''));
}

// 恢复：按标记回原位置（目标已存在同名模板时报错，由前端提示处理）
function restoreTemplate(entryId) {
  const src = resolveRecycleEntry(entryId);
  if (!src || !fs.existsSync(src)) throw new Error('回收站条目不存在');
  const marker = readJsonFile(path.join(src, RECYCLE_MARKER)) || {};
  const origFolder = marker.folder !== undefined ? marker.folder : '';
  const id = safeId(marker.id || path.basename(src));
  const dest = path.join(templateRoot(), origFolder ? safeFolder(origFolder) : '', id);
  if (!dest.startsWith(path.resolve(templateRoot()) + path.sep) && dest !== path.resolve(templateRoot())) {
    throw new Error('非法恢复目标');
  }
  if (fs.existsSync(dest)) throw new Error('目标位置已存在同名模板（' + (origFolder || '根目录') + '/' + id + '），请先处理再恢复');
  fs.mkdirSync(path.dirname(dest), { recursive: true });
  fs.renameSync(src, dest);
  try { fs.rmSync(path.join(dest, RECYCLE_MARKER), { force: true }); } catch { /* ignore */ }
  return { ok: true, id, folder: origFolder };
}

// 彻底删除（不可恢复）
function purgeTemplate(entryId) {
  const target = resolveRecycleEntry(entryId);
  if (!target || target === recycleDir()) return { ok: false, error: '非法条目' };
  if (!fs.existsSync(target)) return { ok: false, error: '回收站条目不存在' };
  fs.rmSync(target, { recursive: true, force: true });
  return { ok: true };
}

// 清空回收站
function emptyRecycleBin() {
  const rc = recycleDir();
  if (!fs.existsSync(rc)) return { ok: true, removed: 0 };
  let removed = 0;
  for (const e of fs.readdirSync(rc, { withFileTypes: true })) {
    const p = path.join(rc, e.name);
    try {
      if (e.isDirectory()) fs.rmSync(p, { recursive: true, force: true });
      else fs.rmSync(p, { force: true });
      removed++;
    } catch { /* 单个失败继续 */ }
  }
  return { ok: true, removed };
}

module.exports = {
  templateRoot, setTemplateRoot, listTemplates, listFolders, invalidateTemplateCache,
  saveTemplate, getTemplate, deleteTemplate, safeId, safeFolder,
  listVersions, getVersion, restoreVersion, setCurrentVersion, deleteVersion,
  recycleDir, RECYCLE_NAME, listRecycleBin, restoreTemplate, purgeTemplate, emptyRecycleBin
};