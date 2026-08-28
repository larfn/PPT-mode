'use strict';

const assert = require('node:assert/strict');
const fs = require('node:fs');
const os = require('node:os');
const path = require('node:path');
const { spawnSync } = require('node:child_process');

const root = path.resolve(__dirname, '..');
const source = path.join(root, '安装.bat');
const tempDir = fs.mkdtempSync(path.join(os.tmpdir(), 'ppt-addin-installer-test-'));
const tempBat = path.join(tempDir, '安装.bat');

try {
  fs.copyFileSync(source, tempBat);
  const bytes = fs.readFileSync(tempBat);
  const sourceText = fs.readFileSync(tempBat, 'utf8');
  const cr = bytes.filter((b) => b === 13).length;
  const lf = bytes.filter((b) => b === 10).length;
  assert.equal(cr, lf, '安装.bat must use CRLF line endings for cmd.exe stability');
  assert.match(sourceText, /taskkill \/F \/T \/IM ppt-ai-addin\.exe/, '安装.bat must stop the stale service process tree before copying files');
  assert.match(sourceText, /scripts\\stop-service\.ps1/, '安装.bat must use the robust PowerShell stop helper before copying files');

  const stopScript = fs.readFileSync(path.join(root, 'scripts', 'stop-service.ps1'), 'utf8');
  assert.match(stopScript, /Stop-Process -Id/, 'stop-service.ps1 must force-stop stale service processes');
  assert.match(stopScript, /Get-NetTCPConnection/, 'stop-service.ps1 must also stop service process by the recorded/listening port');
  assert.match(stopScript, /Stale service is still running/, 'stop-service.ps1 must fail clearly if the old service cannot be stopped');

  const result = spawnSync(
    'cmd.exe',
    ['/d', '/c', 'call 安装.bat --skip-build --no-pause --skip-e2e'],
    { cwd: tempDir, encoding: 'utf8', windowsHide: true }
  );
  const output = `${result.stdout || ''}${result.stderr || ''}`;

  assert.notEqual(result.status, 255, output);
  assert.equal(result.status, 1, output);
  assert.match(output, /release\.json/);
  assert.doesNotMatch(output, /unexpected at this time|不应有/);

  console.log('[installer-test] 安装.bat skip-build 解析检查通过');
} finally {
  fs.rmSync(tempDir, { recursive: true, force: true });
}
