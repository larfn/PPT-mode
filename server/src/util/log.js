// 调试日志工具：写入 %APPDATA%\ppt-ai-addin\ 下，超过上限自动轮转（保留一份 .old），
// 避免 debug-read.log 无限增长占满磁盘。
const fs = require('node:fs');
const os = require('node:os');
const path = require('node:path');

const MAX_LOG_BYTES = 1024 * 1024; // 单文件上限 1MB

function logDir() {
  return path.join(process.env.APPDATA || path.join(os.homedir(), 'AppData', 'Roaming'), 'ppt-ai-addin');
}

// 追加一条 JSON 日志；任何失败都静默（日志不能影响主流程）
function appendLog(file, obj) {
  try {
    const dir = logDir();
    fs.mkdirSync(dir, { recursive: true });
    const p = path.join(dir, file);
    if (fs.existsSync(p) && fs.statSync(p).size > MAX_LOG_BYTES) {
      try { fs.renameSync(p, p + '.old'); } catch { /* 轮转失败则忽略 */ }
    }
    fs.appendFileSync(p, JSON.stringify({ at: new Date().toISOString(), body: obj }) + '\n', 'utf8');
  } catch { /* ignore */ }
}

module.exports = { appendLog, logDir, MAX_LOG_BYTES };
