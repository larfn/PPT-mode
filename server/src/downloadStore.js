// 图片下载：安全下载器 + 本地缓存
// 安全：SSRF 防护（协议/主机/DNS 全解析地址校验，重定向每跳复检）、MIME 白名单、
//       魔数校验（防 content-type 谎报/SVG 脚本）、20MB 大小上限（防内存爆炸）、
//       超时（60s）、原子写入（临时文件→rename，断线不留半文件）、非图片响应拒绝。
// 缓存：key = provider|imageUrl，命中直接复用（避免重复下载同一张图）。
'use strict';
const fs = require('node:fs');
const os = require('node:os');
const path = require('node:path');
const crypto = require('node:crypto');
const {
  ALLOWED_IMAGE_MIME, validateDownloadUrl, sniffImageMime, extFromMime
} = require('./downloadSecurity.js');

const MAX_DOWNLOAD_SIZE = 20 * 1024 * 1024; // 单图 20MB 上限
const DOWNLOAD_TIMEOUT = 60000;             // 单次请求 60s 超时
const MAX_REDIRECTS = 5;

let root = null;
function getDownloadDir() {
  if (root) return root;
  return process.env.PPT_DOWNLOAD_DIR || path.join(os.homedir(), 'Documents', 'PPT下载图库');
}
function setDownloadRoot(dir) { root = dir; }

// 内存任务表（前端轮询期间存在；重启丢失可接受）
const tasks = new Map();
// 下载缓存：key = provider|imageUrl → { filePath, mime, size, savedAt }（不缓存大 base64，命中时读文件重建）
const cache = new Map();
const MAX_CACHE = 200;

function uniqueFileName(dir, ext) {
  const now = new Date();
  const pad = (n) => String(n).padStart(2, '0');
  const base = `PPT图_${now.getFullYear()}${pad(now.getMonth() + 1)}${pad(now.getDate())}_${pad(now.getHours())}${pad(now.getMinutes())}${pad(now.getSeconds())}`;
  let name = `${base}.${ext}`;
  let i = 1;
  while (fs.existsSync(path.join(dir, name))) {
    name = `${base}_${i}.${ext}`;
    i += 1;
  }
  return name;
}

function cacheKeyFor(provider, url) { return (provider || '') + '|' + url; }

// 创建下载任务：立即返回 taskId；provider 用于缓存 key（可选）
function startDownload(url, opts = {}) {
  const id = crypto.randomUUID();
  const task = {
    id, url,
    status: 'downloading',
    received: 0, total: null,
    fileName: '', filePath: '', dataUrl: '', error: '',
    provider: opts.provider || null,
    fromCache: false
  };
  tasks.set(id, task);
  if (tasks.size > 200) {
    const oldest = tasks.keys().next().value;
    if (oldest) tasks.delete(oldest);
  }
  run(task, opts).catch(() => {});
  return id;
}

async function run(task, opts) {
  const url = task.url;
  const cacheKey = cacheKeyFor(opts.provider, url);
  const t0 = process.hrtime.bigint();
  const { record } = require('./perf.js');
  const finish = (status) => {
    record('imageDownload', Number(process.hrtime.bigint() - t0) / 1e6, { status, kb: Math.round((task.received || 0) / 1024), fromCache: task.fromCache || false, provider: opts.provider || null });
  };
  try {
    // 1) 缓存命中：直接复用已下载文件（重建 dataUrl 供前端使用）
    const hit = cache.get(cacheKey);
    if (hit && fs.existsSync(hit.filePath)) {
      const buf = fs.readFileSync(hit.filePath);
      task.status = 'done';
      task.dataUrl = `data:${hit.mime};base64,${buf.toString('base64')}`;
      task.fileName = path.basename(hit.filePath);
      task.filePath = hit.filePath;
      task.received = buf.length;
      task.total = buf.length;
      task.fromCache = true;
      finish('done');
      return;
    }
    // 2) 下载（SSRF 校验 + 手动重定向逐跳复检 + 超时 + 大小上限 + MIME/魔数校验）
    const { buf, mime } = await fetchImageWithGuard(url);
    const ext = extFromMime(mime) || 'jpg';
    const dir = getDownloadDir();
    fs.mkdirSync(dir, { recursive: true });
    const fileName = uniqueFileName(dir, ext);
    const filePath = path.join(dir, fileName);
    // 3) 原子写入：临时文件 → rename（断线/崩溃不留半文件）
    const tmp = filePath + '.tmp-' + process.pid + '-' + Date.now().toString(36);
    fs.writeFileSync(tmp, buf);
    try { fs.renameSync(tmp, filePath); } catch (e) { try { fs.rmSync(tmp, { force: true }); } catch { /* ignore */ } throw e; }
    task.dataUrl = `data:${mime};base64,${buf.toString('base64')}`;
    task.fileName = fileName;
    task.filePath = filePath;
    task.received = buf.length;
    task.total = buf.length;
    task.status = 'done';
    finish('done');
    // 4) 缓存（含文件路径；重复下载直接复用）
    cache.set(cacheKey, { filePath, mime: mime, size: buf.length, savedAt: Date.now() });
    if (cache.size > MAX_CACHE) {
      const k = cache.keys().next().value;
      if (k) cache.delete(k);
    }
  } catch (e) {
    task.status = 'error';
    task.error = e.message || String(e);
    finish('error');
  }
}

// 下载单张图片：SSRF 校验 → 手动跟随重定向（每跳复检）→ 流式读取 + 大小限制
async function fetchImageWithGuard(url) {
  let current = url;
  for (let hop = 0; hop <= MAX_REDIRECTS; hop++) {
    const guard = await validateDownloadUrl(current);
    if (!guard.ok) throw new Error(guard.error);
    const ctrl = new AbortController();
    const timer = setTimeout(() => ctrl.abort(), DOWNLOAD_TIMEOUT);
    let res;
    try {
      res = await fetch(current, {
        headers: { 'User-Agent': 'Mozilla/5.0 ppt-ai-addin' },
        redirect: 'manual', // 手动跟随，逐跳做 SSRF 复检
        signal: ctrl.signal
      });
    } catch (e) {
      if (e && e.name === 'AbortError') throw new Error(`下载超时（超过 ${DOWNLOAD_TIMEOUT / 1000} 秒）`);
      throw new Error('下载失败：' + ((e && e.message) || String(e)));
    } finally {
      clearTimeout(timer);
    }
    if (res.status >= 300 && res.status < 400) {
      const loc = res.headers.get('location');
      if (!loc) throw new Error('重定向缺少 Location');
      try { current = new URL(loc, current).toString(); } catch { throw new Error('非法重定向地址'); }
      continue; // 下一跳重新做 URL 校验（防重定向到内网）
    }
    if (!res.ok) throw new Error(`下载失败（HTTP ${res.status}）`);
    const contentType = (res.headers.get('content-type') || '').split(';')[0].trim().toLowerCase();
    if (!ALLOWED_IMAGE_MIME.has(contentType)) {
      if (contentType === 'image/svg+xml' || contentType === 'image/svg') throw new Error('不支持 SVG 图片（可能包含脚本）');
      throw new Error(`非图片响应（${contentType || '未知类型'}）`);
    }
    // 流式读取 + 大小上限（防恶意超大图片导致内存爆炸）
    const chunks = [];
    let size = 0;
    const body = res.body;
    if (body && typeof body[Symbol.asyncIterator] === 'function') {
      for await (const chunk of body) {
        const buf = Buffer.from(chunk);
        size += buf.length;
        if (size > MAX_DOWNLOAD_SIZE) throw new Error(`图片过大（超过 ${Math.round(MAX_DOWNLOAD_SIZE / 1024 / 1024)}MB 上限）`);
        chunks.push(buf);
      }
    } else {
      const arr = Buffer.from(await res.arrayBuffer());
      size = arr.length;
      if (size > MAX_DOWNLOAD_SIZE) throw new Error(`图片过大（超过 ${Math.round(MAX_DOWNLOAD_SIZE / 1024 / 1024)}MB 上限）`);
      chunks.push(arr);
    }
    const buf = Buffer.concat(chunks);
    if (!buf.length) throw new Error('图片内容为空');
    // 魔数校验：防 content-type 谎报 / SVG 脚本 / 扩展名错误
    const sniffed = sniffImageMime(buf);
    if (!sniffed) throw new Error('不支持的图片格式（仅支持 JPG/PNG/WebP/GIF，SVG 已拒绝）');
    if (sniffed !== contentType && !(sniffed === 'image/jpeg' && contentType === 'image/jpg')) {
      throw new Error('图片内容与声明格式不符（声明 ' + contentType + '，实际 ' + sniffed + '）');
    }
    return { buf, mime: sniffed };
  }
  throw new Error('重定向次数过多');
}

function getTask(id) {
  return tasks.get(id) || null;
}

function ensureDownloadDir() {
  const dir = getDownloadDir();
  fs.mkdirSync(dir, { recursive: true });
  return dir;
}

module.exports = {
  getDownloadDir, setDownloadRoot, startDownload, getTask, ensureDownloadDir,
  _getCache: () => cache, _clearCache: () => cache.clear(),
  _fetchImageWithGuard: fetchImageWithGuard
};
