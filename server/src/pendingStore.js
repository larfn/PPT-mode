// AI 待写队列：MCP/AI 生成的页面先落盘，再由「COM 直接插入」或「任务窗格提示写入」消费。
// 存储：%APPDATA%\ppt-ai-addin\pending\<id>.json（含生成好的 base64）
const fs = require('node:fs');
const os = require('node:os');
const path = require('node:path');

let root = null; // 测试可覆盖

function pendingDir() {
  if (root) return root;
  return process.env.PPT_PENDING_ROOT ||
    path.join(process.env.APPDATA || path.join(os.homedir(), 'AppData', 'Roaming'), 'ppt-ai-addin', 'pending');
}

function setPendingRoot(dir) { root = dir; }

// 清理过期条目：已写入（written=true）超过 24 小时、未写入超过 3 天的删除；损坏文件直接删
// listPending 惰性触发；队列很小，每次全扫成本可忽略
function cleanupPending() {
  const dir = pendingDir();
  if (!fs.existsSync(dir)) return 0;
  const now = Date.now();
  const DAY = 24 * 3600 * 1000;
  let removed = 0;
  for (const f of fs.readdirSync(dir)) {
    if (!f.endsWith('.json')) continue;
    const p = path.join(dir, f);
    try {
      const e = JSON.parse(fs.readFileSync(p, 'utf8'));
      const created = new Date(e.createdAt || 0).getTime();
      if (!Number.isFinite(created) || created <= 0) { fs.rmSync(p, { force: true }); removed++; continue; }
      const maxAge = e.written === true ? DAY : 3 * DAY;
      if (now - created > maxAge) { fs.rmSync(p, { force: true }); removed++; }
    } catch { fs.rmSync(p, { force: true }); removed++; }
  }
  return removed;
}

function listPending() {
  cleanupPending();
  const dir = pendingDir();
  if (!fs.existsSync(dir)) return [];
  return fs.readdirSync(dir)
    .filter((f) => f.endsWith('.json'))
    .map((f) => {
      try { return JSON.parse(fs.readFileSync(path.join(dir, f), 'utf8')); }
      catch { return null; }
    })
    .filter(Boolean)
    .sort((a, b) => (a.createdAt || '').localeCompare(b.createdAt || ''));
}

function addPending(entry) {
  const dir = pendingDir();
  fs.mkdirSync(dir, { recursive: true });
  const id = entry.id || Date.now().toString(36) + Math.random().toString(36).slice(2, 6);
  const full = { id, createdAt: new Date().toISOString(), written: false, ...entry };
  fs.writeFileSync(path.join(dir, id + '.json'), JSON.stringify(full, null, 2), 'utf8');
  return full;
}

function getPending(id) {
  const safe = String(id || '').replace(/[^A-Za-z0-9_-]/g, '');
  if (!safe) return null;
  const p = path.join(pendingDir(), safe + '.json');
  if (!fs.existsSync(p)) return null;
  try { return JSON.parse(fs.readFileSync(p, 'utf8')); } catch { return null; }
}

function markWritten(id) {
  const entry = getPending(id);
  if (!entry) return null;
  entry.written = true;
  entry.writtenAt = new Date().toISOString();
  fs.writeFileSync(path.join(pendingDir(), id + '.json'), JSON.stringify(entry, null, 2), 'utf8');
  return entry;
}

function deletePending(id) {
  const safe = String(id || '').replace(/[^A-Za-z0-9_-]/g, '');
  if (!safe) return;
  const p = path.join(pendingDir(), safe + '.json');
  if (fs.existsSync(p)) fs.rmSync(p, { force: true });
}

function clearAllPending() {
  const dir = pendingDir();
  if (!fs.existsSync(dir)) return 0;
  let removed = 0;
  for (const f of fs.readdirSync(dir)) {
    if (!f.endsWith('.json')) continue;
    fs.rmSync(path.join(dir, f), { force: true });
    removed++;
  }
  return removed;
}

module.exports = { pendingDir, setPendingRoot, listPending, addPending, getPending, markWritten, deletePending, cleanupPending, clearAllPending };
