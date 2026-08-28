// 当前演示文稿上下文（PPT Context）：通过 PowerShell COM 读取正在运行的 PowerPoint，
// 返回结构化、裁剪过的 JSON —— 绝不读取/传输整个 PPTX 二进制。
//
// 链路：MCP/HTTP → 本模块 → context.ps1（GetActiveObject 连运行中的 PowerPoint）
//  - 不依赖任务窗格（Office.js 只能在任务窗格运行，MCP 是独立进程）
//  - 只读需要的信息；文本/形状数量均裁剪（context.ps1 内 MAX_TEXT/MAX_SHAPES 限制）
//  - 稳定 ID：slide 用 SlideID（文档内持久，移动/插入后不变），shape 用 Shape.Id + Name
'use strict';
const { spawn } = require('node:child_process');
const fs = require('node:fs');
const os = require('node:os');
const path = require('node:path');

// pkg 打包后 __dirname 指向 exe 内虚拟路径（C:\\snapshot\\...），powershell -File 无法执行虚拟路径；
// 因此若脚本文件不在磁盘上，先从打包资源读内容解压到临时目录再执行。
const SCRIPT = path.join(__dirname, '..', 'mcp', 'context.ps1');
function resolveScript() {
  // 关键：pkg 打包后 __dirname 是虚拟路径（C:\\snapshot\\...），fs.existsSync 对打包资产返回 true，
  // 但 powershell.exe（外部进程）读不到虚拟路径 → 必须解压到临时目录。
  // 判定打包环境用 process.pkg（源码 node 运行无此字段），不能靠 existsSync。
  const isPkg = typeof process.pkg !== 'undefined';
  if (!isPkg && fs.existsSync(SCRIPT)) return SCRIPT; // 开发环境（node 直接跑）直接可用
  try {
    const content = fs.readFileSync(SCRIPT, 'utf8'); // pkg assets 可读虚拟文件
    const body = content.charCodeAt(0) === 0xFEFF ? content.slice(1) : content; // 去重：资产自带 BOM 时不重复加
    const tmp = path.join(os.tmpdir(), 'ppt-ctx-' + process.pid + '-' + Date.now() + '.ps1');
    fs.writeFileSync(tmp, '\uFEFF' + body, 'utf8');
    return tmp;
  } catch {
    return SCRIPT; // 兜底：交给 spawn 报错
  }
}

// 测试可注入的 runner：(args, outFile) => Promise<object>
let psRunner = null;
function setPsRunner(fn) { psRunner = fn; }

function runPs(args) {
  const outFile = path.join(os.tmpdir(), 'ppt-ctx-' + process.pid + '-' + Date.now() + '-' + Math.random().toString(36).slice(2, 8) + '.json');
  const fullArgs = ['-NoProfile', '-ExecutionPolicy', 'Bypass', '-File', resolveScript(), '-OutFile', outFile, ...args];
  if (psRunner) {
    return Promise.resolve(psRunner(fullArgs, outFile)).then((data) => sanitize(data));
  }
  return new Promise((resolve) => {
    const child = spawn('powershell.exe', fullArgs, { windowsHide: true });
    let err = '';
    child.stderr.on('data', (d) => { err += d; });
    child.on('error', (e) => resolve({ ok: false, error: '无法启动 PowerShell（读取 PPT 上下文失败）：' + e.message }));
    child.on('close', (code) => {
      let data = null;
      try {
        // PS 5.1 Out-File utf8 会带 BOM，需剥离
        const raw = fs.readFileSync(outFile, 'utf8').replace(/^\uFEFF/, '');
        data = JSON.parse(raw);
      } catch { /* 文件不存在或非 JSON */ }
      try { fs.rmSync(outFile, { force: true }); } catch { /* ignore */ }
      if (code === 0 && data) resolve(sanitize(data));
      else resolve({ ok: false, error: (data && data.error) || (err || ('读取 PPT 上下文失败（exit ' + code + '）')).trim().slice(0, 500) });
    });
  });
}

// 清洗：PS 5.1 ConvertTo-Json 会把嵌套空数组序列化为 null；统一归一为 [] 并去掉 null 噪音
function sanitize(v, depth = 0) {
  if (depth > 20) return undefined;
  if (Array.isArray(v)) {
    const out = v.map((x) => sanitize(x, depth + 1)).filter((x) => x !== null && x !== undefined);
    return out;
  }
  if (v && typeof v === 'object') {
    const out = {};
    for (const [k, x] of Object.entries(v)) {
      if (x === null || x === undefined) continue; // 空数组(null)与空值一律省略
      out[k] = sanitize(x, depth + 1);
    }
    return out;
  }
  return v;
}

async function getPresentationContext() { return runPs(['-Action', 'presentation']); }
async function getCurrentDocPath() { return runPs(['-Action', 'doc-path']); }
async function getCurrentSlide() { return runPs(['-Action', 'current-slide']); }
async function getSlide({ index, id } = {}) {
  const args = ['-Action', 'slide'];
  if (id !== undefined && id !== null && id !== '') args.push('-Id', String(id));
  else if (index !== undefined && index !== null && index !== '') args.push('-Index', String(index));
  else return { ok: false, error: 'slide index 或 slideId 必填' };
  return runPs(args);
}
async function inspectSlide({ index, id } = {}) {
  const args = ['-Action', 'inspect'];
  if (id !== undefined && id !== null && id !== '') args.push('-Id', String(id));
  else if (index !== undefined && index !== null && index !== '') args.push('-Index', String(index));
  return runPs(args);
}

module.exports = { getPresentationContext, getCurrentSlide, getSlide, inspectSlide, getCurrentDocPath, setPsRunner, sanitize };
