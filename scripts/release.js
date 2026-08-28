// 发布流水线：统一版本号 + 测试 + 前端构建 + exe 构建 + release 清单
//
// 用法（项目根目录）：
//   node scripts/release.js              # 完整发布：版本号+测试+前端+exe+manifest+清单
//   node scripts/release.js --skip-tests # 跳过测试（紧急构建）
//   node scripts/release.js --skip-build # 只重算版本号+写 manifest/清单（复用现有产物）
//
// 版本号规则：YYYY.MM.DD.NN（当天第 NN 次构建）。单一版本源：
//   1. 本脚本计算版本号并写入 server/src/version.js（打进 exe）与 release.json（供 vite 注入前端）；
//   2. addin 构建时 vite 从 release.json 读取版本注入 __APP_VERSION__ / __APP_BUILT_AT__；
//   3. manifest.xml 的 <Version> 保持固定（2026.08.24.02，2026-08-24 起不再改写）——
//      Office 会把「清单版本 ≠ 已信任版本」的加载项判为无效，每次部署改版本号
//      正是「加载项无效、需手动重新选择」的根因（用户 2026-08-24 指定修复）；
//   4. release.json 汇总产物大小/哈希，供 安装.bat 的 Verify 步骤比对。
'use strict';
const fs = require('node:fs');
const path = require('node:path');
const crypto = require('node:crypto');
const { spawnSync } = require('node:child_process');

const ROOT = path.resolve(__dirname, '..');
const SERVER_DIR = path.join(ROOT, 'server');
const ADDIN_DIR = path.join(ROOT, 'addin');
const RELEASE_FILE = path.join(ROOT, 'release.json');
const VERSION_JS = path.join(SERVER_DIR, 'src', 'version.js');
const MANIFEST_XML = path.join(ROOT, 'manifest.xml');
const EXE = path.join(ROOT, 'dist-exe', 'ppt-ai-addin.exe');
const API_VERSION = '1'; // API 契约版本：破坏性变更（客户端需同步改）时手动 +1

function run(cmdStr, opts = {}) {
  console.log('\n>> ' + cmdStr);
  const r = spawnSync(cmdStr, { cwd: opts.cwd || ROOT, shell: true, stdio: 'inherit', encoding: 'utf8' });
  if (r.error) { console.error('[release] 无法执行：' + r.error.message); process.exit(1); }
  if (r.status !== 0) {
    console.error('[release] 命令失败（退出码 ' + r.status + '）：' + cmdStr);
    process.exit(r.status || 1);
  }
}

function readRelease() {
  try { return JSON.parse(fs.readFileSync(RELEASE_FILE, 'utf8')); } catch { return null; }
}
function writeRelease(data) { fs.writeFileSync(RELEASE_FILE, JSON.stringify(data, null, 2) + '\n'); }

function todayBase() {
  const d = new Date();
  const p = (n) => String(n).padStart(2, '0');
  return d.getFullYear() + '.' + p(d.getMonth() + 1) + '.' + p(d.getDate());
}
// 当天第一次构建从 .01 起，同一天每构建一次序号 +1
function nextVersion(old) {
  const base = todayBase();
  if (old && typeof old === 'string' && old.startsWith(base + '.')) {
    const seq = parseInt(old.slice(base.length + 1), 10) || 0;
    return base + '.' + String(seq + 1).padStart(2, '0');
  }
  return base + '.01';
}

function sha256(file) {
  return crypto.createHash('sha256').update(fs.readFileSync(file)).digest('hex');
}

function collectArtifacts() {
  const out = {};
  if (fs.existsSync(EXE)) {
    out.exe = { path: 'dist-exe/ppt-ai-addin.exe', size: fs.statSync(EXE).size, sha256: sha256(EXE) };
  }
  const assets = path.join(ADDIN_DIR, 'dist', 'assets');
  if (fs.existsSync(assets)) {
    const js = fs.readdirSync(assets).filter((f) => /^index-.*\.js$/.test(f)).sort();
    out.frontend = {
      path: 'addin/dist',
      mainJs: js[0] || null,
      mainJsSize: js[0] ? fs.statSync(path.join(assets, js[0])).size : 0
    };
  }
  if (fs.existsSync(MANIFEST_XML)) out.manifest = { path: 'manifest.xml', sha256: sha256(MANIFEST_XML) };
  return out;
}

function main() {
  const flags = new Set(process.argv.slice(2));
  const skipTests = flags.has('--skip-tests');
  const skipBuild = flags.has('--skip-build');
  const skipE2E = flags.has('--skip-e2e');

  const old = readRelease();
  const version = nextVersion(old && old.version);
  const builtAt = new Date().toISOString();
  console.log('══════════════════════════════════════════════');
  console.log('  模板助手发布流水线');
  console.log('  版本：' + version + '  构建时间：' + builtAt + (skipTests ? '  [跳过测试]' : '') + (skipE2E ? '  [跳过E2E]' : '') + (skipBuild ? '  [跳过构建]' : ''));
  console.log('══════════════════════════════════════════════');

  // 1. 生成后端版本模块（打进 exe；开发/测试默认 0.0.0-dev 会被覆盖）
  fs.writeFileSync(VERSION_JS,
    "// 自动生成：由 scripts/release.js 每次发布时改写，请勿手改。\n" +
    "'use strict';\n" +
    "module.exports = { VERSION: " + JSON.stringify(version) + ", API_VERSION: " + JSON.stringify(API_VERSION) + ", BUILT_AT: " + JSON.stringify(builtAt) + " };\n");
  console.log('[release] server/src/version.js → ' + version);

  // 2. 先写 release.json（含版本），addin 的 vite 构建从这里注入版本
  writeRelease({ version, apiVersion: API_VERSION, builtAt, name: 'ppt-ai-addin', artifacts: {}, tests: {} });

  // 3. 后端测试
  if (!skipTests) {
    run('npm run test:installer');
    run('npm test', { cwd: SERVER_DIR });
    run('npm test', { cwd: ADDIN_DIR });
  }

  // 4. 黄金路径 E2E（发布门禁）：6 条真实用户路径，文件级必跑 + PowerPoint COM 可用时自动打开检查。
  //    （--skip-tests 不影响 E2E；安装.bat 默认跳过 npm test 但 E2E 仍会跑，约 10 秒）
  if (!skipE2E) {
    run('node e2e/golden-path.js', { cwd: SERVER_DIR });
  }

  // 5. 前端：类型检查 + 构建
  if (!skipBuild) {
    run('node node_modules/typescript/bin/tsc --noEmit', { cwd: ADDIN_DIR });
    run('npm run build', { cwd: ADDIN_DIR });
  }

  // 6. 后端 exe 打包
  if (!skipBuild) {
    run('npm run build:exe', { cwd: SERVER_DIR });
  }

  // 7. manifest.xml 版本固定（不再改写，见文件头注释：避免 Office 信任记录失效）
  // 8. 汇总产物信息 → release.json
  const artifacts = collectArtifacts();
  writeRelease({
    version,
    apiVersion: API_VERSION,
    builtAt,
    name: 'ppt-ai-addin',
    artifacts,
    tests: {
      e2e: skipE2E ? 'skipped' : 'passed',
      server: skipTests ? 'skipped' : 'passed',
      addin: skipTests ? 'skipped' : 'passed',
      tsc: skipBuild ? 'skipped' : 'passed',
      addinBuild: skipBuild ? 'skipped' : 'passed',
      exeBuild: skipBuild ? 'skipped' : 'passed'
    }
  });

  // 9. 摘要
  console.log('\n──────────────────────────────────────────────');
  console.log('  发布完成：' + version);
  if (artifacts.exe) console.log('  exe       ' + (artifacts.exe.size / 1048576).toFixed(1) + ' MB  sha256=' + artifacts.exe.sha256.slice(0, 16) + '…');
  if (artifacts.frontend) console.log('  前端主包  ' + (artifacts.frontend.mainJsSize / 1024).toFixed(1) + ' kB  ' + (artifacts.frontend.mainJs || ''));
  if (artifacts.manifest) console.log('  manifest  sha256=' + artifacts.manifest.sha256.slice(0, 16) + '…');
  console.log('  release.json 已生成，可运行 安装.bat（Verify 步骤会比对运行版本）');
  console.log('──────────────────────────────────────────────');
}

main();
