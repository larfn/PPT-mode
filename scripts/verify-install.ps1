# 模板助手安装自检（安装.bat 第 6 步调用，也可单独运行）
# 职责：启动后核对「实际运行版本」与「发布清单 release.json」是否一致，
#       并检查 Runtime Token / 后端连接 / 前端页面 / MCP 状态。
# 退出码：0 = 全部通过；1 = 存在失败项。
param()
$ErrorActionPreference = 'Stop'
$host.UI.RawUI.ForegroundColor = 'Gray'

$rtFile = Join-Path $env:APPDATA 'ppt-ai-addin\runtime.json'
$relFile = Join-Path $env:LOCALAPPDATA 'PPT-AI-Addin\release.json'

$pass = 0; $fail = 0
function Check([string]$name, [bool]$ok, [string]$detail) {
  if ($ok) { $script:pass++; Write-Host ("  [PASS] " + $name + " — " + $detail) -ForegroundColor Green }
  else     { $script:fail++; Write-Host ("  [FAIL] " + $name + " — " + $detail) -ForegroundColor Red }
}
# 信息项：不影响安装成败（如 MCP 未运行属正常）
function Info([string]$name, [string]$detail) {
  Write-Host ("  [INFO] " + $name + " — " + $detail) -ForegroundColor Yellow
}

Write-Host "========== 模板助手安装自检 ==========" -ForegroundColor Cyan

# 1. runtime.json（实际端口 + 一次性 token）
$rt = $null; $port = 3788; $token = ''
if (Test-Path -LiteralPath $rtFile) {
  $rt = Get-Content -Raw -LiteralPath $rtFile | ConvertFrom-Json
  Check 'Runtime Token' -ok ([bool]$rt.token) -detail ($(if ($rt.token) { '已生成（64 hex）' } else { '缺失' }))
  $port = [int]$rt.port
} else {
  Check 'Runtime Token' -ok $false -detail ('runtime.json 未生成：' + $rtFile)
}

# 2. 后端连接 + 版本（/api/version 需要 token）
$version = $null
$headers = @{}
if ($rt -and $rt.token) { $headers['X-Auth-Token'] = [string]$rt.token }
try {
  $version = Invoke-RestMethod -Uri ("http://127.0.0.1:{0}/api/version" -f $port) -Headers $headers -TimeoutSec 5
  Check '后端连接' -ok $true -detail ("端口 {0}，版本 {1}" -f $port, $version.appVersion)
} catch {
  Check '后端连接' -ok $false -detail ("端口 {0}：{1}" -f $port, $_.Exception.Message)
}

# 3. 发布清单比对（关键：代码更新了但 exe 是旧的 → 这里直接暴露）
if (Test-Path -LiteralPath $relFile) {
  $rel = Get-Content -Raw -LiteralPath $relFile | ConvertFrom-Json
  Check '发布清单存在' -ok $true -detail ("版本 {0}，构建于 {1}" -f $rel.version, $rel.builtAt)
  if ($version) {
    Check '运行版本=发布版本' -ok ($version.appVersion -eq $rel.version) -detail ("实际运行 {0} vs 清单 {1}" -f $version.appVersion, $rel.version)
    Check '前端资源' -ok ($null -ne $version.frontend.sizeKB) -detail ($(if ($null -ne $version.frontend.sizeKB) { "{0} kB" -f $version.frontend.sizeKB } else { '未找到 index-*.js' }))
  } else {
    Check '运行版本=发布版本' -ok $false -detail '无法获取后端版本（连接失败）'
  }
} else {
  Check '发布清单存在' -ok $false -detail ("未找到 " + $relFile)
}

# 4. MCP 状态（信息项：MCP 是 stdio 服务，仅 AI 客户端会话期间在线，未运行属正常）
if ($version) {
  if ($version.mcp.running) { Info 'MCP 状态' ("运行中（PID " + $version.mcp.pid + "）") }
  else { Info 'MCP 状态' '未运行（属正常：AI 客户端连接时自动在线，或运行 启动MCP.bat）' }
}

# 5. 前端页面可达
try {
  $idx = Invoke-WebRequest -Uri ("http://127.0.0.1:{0}/index.html" -f $port) -TimeoutSec 5 -UseBasicParsing
  Check '前端页面' -ok ($idx.StatusCode -eq 200) -detail ("HTTP {0}" -f $idx.StatusCode)
} catch {
  Check '前端页面' -ok $false -detail $_.Exception.Message
}

Write-Host ""
Write-Host ("自检结果：{0} 项通过，{1} 项失败" -f $pass, $fail) -ForegroundColor $(if ($fail -eq 0) { 'Green' } else { 'Red' })
Write-Host "提示：MCP 未运行属正常（仅 AI 客户端会话期间在线）；关键失败（后端连接/版本不一致）请重新运行 安装.bat。"
if ($fail -gt 0) { exit 1 } else { exit 0 }