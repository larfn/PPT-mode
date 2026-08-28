// 路径常量（独立模块，避免 config.js ↔ security.js 循环依赖）
'use strict';
const os = require('node:os');
const path = require('node:path');

const DATA_DIR = path.join(process.env.APPDATA || path.join(os.homedir(), '.config'), 'ppt-ai-addin');
const CONFIG_FILE = path.join(DATA_DIR, 'config.json');

module.exports = { DATA_DIR, CONFIG_FILE };
