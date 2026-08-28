// 安全加固（P2-E）：Windows DPAPI 密钥加密 + 一次性访问 token（runtime.json）
//
// 职责：
//   1. DPAPI（Data Protection API，CurrentUser 作用域）对 API Key 加密/解密。
//      无原生依赖：复用项目已有的 powershell.exe 通道调用
//      System.Security.Cryptography.ProtectedData（Windows 专属；非 Windows
//      环境自动回退明文，保证可移植与测试可跑）。采用 spawnSync 同步执行，
//      保持 config.js 的同步 API 不变（加密只发生在保存配置/首次读取时）。
//   2. runtime.json：后端每次启动生成「一次性 token」并记录实际监听端口，
//      写进 %APPDATA%/ppt-ai-addin/runtime.json。MCP server 与前端（/api/runtime）
//      从同一文件获取 token，实现「随机端口 + 一次性 token，前端与 MCP 共用」。
//      仅监听 127.0.0.1 + token 双重收窄本机暴露面。
'use strict';
const fs = require('node:fs');
const path = require('node:path');
const crypto = require('node:crypto');
const { spawnSync } = require('node:child_process');
const { DATA_DIR } = require('./paths.js');

const RUNTIME_FILE = path.join(DATA_DIR, 'runtime.json');

// ---------- DPAPI（PowerShell ProtectedData） ----------
const DPAPI_PREFIX = 'dpapi:';

function isWindows() { return process.platform === 'win32'; }

// 执行 PowerShell 单行命令（同步），返回 stdout（trim 后）；失败返回 null。
function runPowershell(script) {
  let r;
  try {
    r = spawnSync('powershell.exe', ['-NoProfile', '-NonInteractive', '-ExecutionPolicy', 'Bypass', '-Command', script], {
      windowsHide: true,
      encoding: 'utf8',
      timeout: 15000
    });
  } catch { return null; }
  if (r.status !== 0 || r.error) return null;
  return String(r.stdout || '').trim();
}

function protectScript(b64) {
  return "Add-Type -AssemblyName System.Security; [Convert]::ToBase64String([System.Security.Cryptography.ProtectedData]::Protect([Convert]::FromBase64String('" + b64 + "'), $null, [System.Security.Cryptography.DataProtectionScope]::CurrentUser))";
}
function unprotectScript(b64) {
  return "Add-Type -AssemblyName System.Security; [Convert]::ToBase64String([System.Security.Cryptography.ProtectedData]::Unprotect([Convert]::FromBase64String('" + b64 + "'), $null, [System.Security.Cryptography.DataProtectionScope]::CurrentUser))";
}

// 加密明文 → 'dpapi:<base64>'；非 Windows / 已是 dpapi: / 空值 → 原样（幂等）
function protectSecret(plain) {
  if (!plain || !isWindows() || String(plain).startsWith(DPAPI_PREFIX)) return plain;
  const b64 = Buffer.from(String(plain), 'utf8').toString('base64');
  const out = runPowershell(protectScript(b64));
  if (!out) { console.warn('[ppt-ai] DPAPI protect 失败，该 Key 将以明文保存（仅本机文件，风险可接受）'); return plain; }
  return DPAPI_PREFIX + out;
}

// 解密 'dpapi:<base64>' → 明文；解密失败（如换 Windows 用户）返回空串并告警
function unprotectSecret(enc) {
  if (!enc) return enc;
  if (typeof enc !== 'string' || !enc.startsWith(DPAPI_PREFIX)) return enc; // 旧版明文/空值透传
  const b64 = enc.slice(DPAPI_PREFIX.length);
  const out = runPowershell(unprotectScript(b64));
  if (!out) {
    console.warn('[ppt-ai] DPAPI unprotect 失败（可能是 Windows 用户/系统环境变化），请重新填写 API Key');
    return '';
  }
  return Buffer.from(out, 'base64').toString('utf8');
}

// ---------- runtime.json（一次性 token + 实际端口） ----------
let runtimeCache = null; // { mtimeMs, data }
function loadRuntime() {
  try {
    const st = fs.statSync(RUNTIME_FILE);
    if (runtimeCache && runtimeCache.mtimeMs === st.mtimeMs) return runtimeCache.data;
    const data = JSON.parse(fs.readFileSync(RUNTIME_FILE, 'utf8'));
    runtimeCache = { mtimeMs: st.mtimeMs, data };
    return data;
  } catch {
    runtimeCache = null;
    return null;
  }
}
function saveRuntime(data) {
  fs.mkdirSync(path.dirname(RUNTIME_FILE), { recursive: true });
  try { fs.writeFileSync(RUNTIME_FILE, JSON.stringify(data, null, 2), { encoding: 'utf8', mode: 0o600 }); }
  catch { fs.writeFileSync(RUNTIME_FILE, JSON.stringify(data, null, 2), 'utf8'); }
  runtimeCache = null; // 强制下次重读
}

// 生成/刷新一次性 token（每次后端启动换新，旧 token 立即失效）
function ensureRuntimeToken() {
  const token = crypto.randomBytes(32).toString('hex');
  const prev = loadRuntime() || {};
  saveRuntime({ ...prev, token, startedAt: new Date().toISOString() });
  return token;
}
function getToken() {
  const rt = loadRuntime();
  return rt && rt.token ? rt.token : '';
}
// 记录实际监听端口（listen 成功回调里调用）
function recordRuntimePort(port) {
  const prev = loadRuntime() || {};
  saveRuntime({ ...prev, port });
}
function getRuntimePort() {
  const rt = loadRuntime();
  return rt && Number.isInteger(rt.port) ? rt.port : null;
}

// auth 是否启用：显式 PPT_AUTH=1 强制启用（供测试/手动加固）；
// 否则仅在「本进程入口是 server/src/index.js（node src/index.js / 打包 exe）」
// 且 runtime.json 已有 token 时启用 —— 测试进程（node --test require 本模块）入口不是
// index.js，自动不启用，现有测试零改动可跑。
function isAuthEnabled() {
  if (process.env.PPT_AUTH === '1') return true;
  const mainFile = require.main && require.main.filename;
  if (mainFile && path.resolve(mainFile) === path.resolve(__dirname, 'index.js')) {
    return !!getToken();
  }
  return false;
}

module.exports = {
  RUNTIME_FILE, DPAPI_PREFIX,
  protectSecret, unprotectSecret, isWindows,
  loadRuntime, saveRuntime, ensureRuntimeToken, getToken, recordRuntimePort, getRuntimePort,
  isAuthEnabled
};