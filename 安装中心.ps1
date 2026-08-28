# ============================================================
#  模板助手 · 安装向导（桌面版）
#  双击 安装中心.exe（或 安装中心.bat）打开安装向导：
#     欢迎 → 环境检查 → 安装选项 → 正在安装 → 完成
#  内部调用 安装.bat / 卸载.bat 完成任务。
# ============================================================
param([switch]$SelfTest, [switch]$GuiTest)
$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
[System.Windows.Forms.Application]::EnableVisualStyles()

# （代码页探测改用 [Console]::OutputEncoding.CodePage，见 Start-BatJob）

$root = $PSScriptRoot
if (-not $root) { $root = (Get-Location).Path }
$batInstall   = Join-Path $root '安装.bat'
$batUninstall = Join-Path $root '卸载.bat'
if (-not (Test-Path -LiteralPath $batInstall)) {
  [System.Windows.Forms.MessageBox]::Show('未找到 安装.bat，请把「安装中心.bat / 安装中心.ps1」放在项目根目录。', '模板助手 · 安装向导', 'OK', 'Error') | Out-Null
  exit 1
}

# 防止重复打开多个向导窗口
$script:mutex = New-Object System.Threading.Mutex($false, 'Local\PPT-AI-Install-Wizard')
if (-not $script:mutex.WaitOne(0)) {
  [System.Windows.Forms.MessageBox]::Show('安装向导已在运行。', '模板助手 · 安装向导', 'OK', 'Information') | Out-Null
  exit 0
}

$script:crlf = [string][char]13 + [string][char]10
$script:nl   = [string][char]10
$script:mode = 'install'        # install | uninstall
$script:step = 0
$script:envOK = $false
$script:job = $null
$script:cancelled = $false
$script:lastDone = $null
$script:logLines = New-Object 'System.Collections.Generic.List[string]'
$script:renderedCount = 0
$script:pendingBytes = New-Object byte[] 0

# ---- 应用图标（优先加载同目录 安装中心.ico）----
$script:appIcon = $null
$icoPath = Join-Path $root '安装中心.ico'
if (Test-Path -LiteralPath $icoPath) { try { $script:appIcon = New-Object System.Drawing.Icon($icoPath) } catch { $script:appIcon = $null } }

function Get-NodeInfo {
  $node = Get-Command node -ErrorAction SilentlyContinue
  $ver = ''
  if ($node) { try { $ver = [string](& $node.Source --version 2>$null | Select-Object -First 1) } catch {} }
  $npm = Get-Command npm -ErrorAction SilentlyContinue
  $npmVer = ''
  if ($npm) { try { $npmVer = [string](& $npm.Source --version 2>$null | Select-Object -First 1) } catch {} }
  return [pscustomobject]@{ Present = [bool]$node; Version = $ver; Npm = $npmVer }
}

function Get-EnvCheck {
  $node = Get-NodeInfo
  $relFile = Join-Path $root 'release.json'
  $rel = $null
  if (Test-Path -LiteralPath $relFile) { try { $rel = Get-Content -Raw -LiteralPath $relFile | ConvertFrom-Json } catch {} }
  $installDir = Join-Path $env:LOCALAPPDATA 'PPT-AI-Addin'
  $irel = $null
  $irf = Join-Path $installDir 'release.json'
  if (Test-Path -LiteralPath $irf) { try { $irel = Get-Content -Raw -LiteralPath $irf | ConvertFrom-Json } catch {} }
  $proc = Get-Process ppt-ai-addin -ErrorAction SilentlyContinue | Select-Object -First 1
  $rt = $null
  $rtFile = Join-Path $env:APPDATA 'ppt-ai-addin\runtime.json'
  if (Test-Path -LiteralPath $rtFile) { try { $rt = Get-Content -Raw -LiteralPath $rtFile | ConvertFrom-Json } catch {} }
  $port = 3788; $svcVer = ''; $mcpRun = $false
  if ($rt) {
    $port = [int]$rt.port
    if ($rt.token) {
      try {
        $v = Invoke-RestMethod -Uri ("http://127.0.0.1:{0}/api/version" -f $port) -Headers @{ 'X-Auth-Token' = [string]$rt.token } -TimeoutSec 2
        $svcVer = [string]$v.appVersion
        $mcpRun = [bool]$v.mcp.running
      } catch {}
    }
  }
  $regOK = $false
  $manifestPath = Join-Path $installDir 'manifest.xml'
  if (Test-Path -LiteralPath $manifestPath) {
    try {
      [xml]$m = Get-Content -Raw -Encoding UTF8 -LiteralPath $manifestPath
      $id = [string]$m.OfficeApp.Id
      $key = Get-ItemProperty -Path 'HKCU:\SOFTWARE\Microsoft\Office\16.0\Wef\Developer' -Name $id -ErrorAction SilentlyContinue
      $regOK = ($null -ne $key -and [string]$key.$id -eq $manifestPath)
    } catch {}
  }
  return [pscustomobject]@{
    NodePresent = [bool]$node.Present; NodeVersion = [string]$node.Version; NpmVersion = [string]$node.Npm
    RelPresent  = ($null -ne $rel);    RelVersion  = $(if ($rel) {[string]$rel.version} else {''}); RelBuilt = $(if ($rel) {[string]$rel.builtAt} else {''})
    InstPresent = ($null -ne $irel);   InstVersion = $(if ($irel) {[string]$irel.version} else {''})
    RegOK = $regOK
    SvcProcess = [bool]$proc; SvcVersion = $svcVer; SvcPort = $port; Mcp = $mcpRun
  }
}

function Start-BatJob([string]$kind, [string]$batFile, [string]$args) {
  if ($script:job -and -not $script:job.Proc.HasExited) { return $false }
  # 清理上一个任务的临时 bat 副本
  if ($script:lastTmpBat -and (Test-Path -LiteralPath $script:lastTmpBat)) { Remove-Item -LiteralPath $script:lastTmpBat -Force -ErrorAction SilentlyContinue }
  $script:lastTmpBat = $null
  $log = Join-Path $env:TEMP ('ppt-addin-ui-' + [guid]::NewGuid().ToString('N') + '.log')
  # cmd 按「启动时控制台代码页」读取 bat（有控制台则继承父控制台代码页；无控制台则新建控制台用 OEM 代码页）。
  # 把 bat（UTF-8 中文）转成该代码页的临时副本放回项目目录，避免 UTF-8 被按 936 错位解析导致乱码/参数失效。
  $runBat = $batFile
  $tmpBat = $null
  try {
    $cp = [int][Console]::OutputEncoding.CodePage
    if ($cp -le 0) { $cp = 936 }
    $content = [System.IO.File]::ReadAllText($batFile, (New-Object System.Text.UTF8Encoding($false)))
    $tmpBat = Join-Path $root ('_tmp_' + [guid]::NewGuid().ToString('N') + '.bat')
    if ([int]$cp -eq 65001) {
      # 控制台已是 65001：UTF-8 bat + chcp 全程一致，直接写 UTF-8 副本
      [System.IO.File]::WriteAllText($tmpBat, $content, (New-Object System.Text.UTF8Encoding($false)))
    } else {
      # 控制台是其他代码页（如 936）：转成该代码页，并去掉 chcp 行（chcp 会把后续行改成别的解码，破坏中文）
      $content = [regex]::Replace($content, '(?im)^@chcp[^
]*?
', '')
      [System.IO.File]::WriteAllText($tmpBat, $content, [System.Text.Encoding]::GetEncoding([int]$cp))
    }
    $runBat = $tmpBat
    $script:lastTmpBat = $tmpBat
  } catch { $runBat = $batFile }
  $psi = New-Object System.Diagnostics.ProcessStartInfo
  $psi.FileName = 'cmd.exe'
  $psi.Arguments = '/d /c ""' + $runBat + '" ' + $args + ' < nul > "' + $log + '" 2>&1"'
  $psi.WorkingDirectory = $root
  $psi.UseShellExecute = $false
  $psi.CreateNoWindow = $false
  $psi.WindowStyle = 'Hidden'
  $p = New-Object System.Diagnostics.Process
  $p.StartInfo = $psi
  $null = $p.Start()
  $script:job = [pscustomobject]@{ Proc = $p; LogPath = $log; LastLen = 0; Kind = $kind; Start = (Get-Date); Step = 0; TotalSteps = 0; StepLabel = '' }
  $script:cancelled = $false
  $script:lastDone = $null
  $script:logLines.Clear()
  $script:renderedCount = 0
  $script:pendingBytes = New-Object byte[] 0
  return $true
}
function Add-ParsedLine([string]$t) {
  $t = $t.TrimEnd()
  if ($t -eq '') { return }
  $script:logLines.Add($t)
  $m = [regex]::Match($t, '^\[([0-9]+)/([0-9]+)\]\s*(.*)$')
  if ($m.Success) {
    $n = [int]$m.Groups[1].Value
    $tot = [int]$m.Groups[2].Value
    if ($tot -ge 1 -and $tot -le 12 -and $n -ge 0 -and $n -le $tot) {
      $script:job.Step = $n
      $script:job.TotalSteps = $tot
      $desc = $m.Groups[3].Value
      $script:job.StepLabel = $(if ($desc) { $desc } else { ('步骤 ' + $n + ' / ' + $tot) })
    }
  }
}

function Decode-Smart([byte[]]$bytes) {
  # 逐行智能解码：先试严格 UTF-8（node/构建输出），失败回退 GBK（cmd echo 输出）
  try { return (New-Object System.Text.UTF8Encoding($false, $true)).GetString($bytes) }
  catch {
    try { return [System.Text.Encoding]::GetEncoding(936).GetString($bytes) } catch { return '' }
  }
}

function Sync-JobLog([switch]$FlushPending) {
  $j = $script:job
  if (-not $j) { return }
  if (-not (Test-Path -LiteralPath $j.LogPath)) { return }
  try {
    # 用 FileShare.ReadWrite 打开：cmd 写日志期间其他进程可读（ReadAllBytes 默认 FileShare.Read 会被拒）
    $fs = New-Object System.IO.FileStream($j.LogPath, [System.IO.FileMode]::Open, [System.IO.FileAccess]::Read, [System.IO.FileShare]::ReadWrite)
    try {
      $fsLen = $fs.Length
      if ($fsLen -le $j.LastLen) { return }
      $null = $fs.Seek($j.LastLen, [System.IO.SeekOrigin]::Begin)
      $toRead = [int]($fsLen - $j.LastLen)
      $bytes = New-Object byte[] $toRead
      $readN = 0
      while ($readN -lt $toRead) {
        $n = $fs.Read($bytes, $readN, $toRead - $readN)
        if ($n -le 0) { break }
        $readN += $n
      }
      $j.LastLen = $j.LastLen + $readN
    } finally { $fs.Close() }
    if ($readN -le 0) { return }
    # 拼接上次未完成的行字节
    $all = New-Object byte[] ($script:pendingBytes.Length + $readN)
    if ($script:pendingBytes.Length -gt 0) { [Array]::Copy($script:pendingBytes, 0, $all, 0, $script:pendingBytes.Length) }
    [Array]::Copy($bytes, 0, $all, $script:pendingBytes.Length, $readN)
    # 按 0x0A 切行，每行独立解码（日志可能是 GBK 与 UTF-8 混合）
    $start = 0
    for ($i = 0; $i -lt $all.Length; $i++) {
      if ($all[$i] -eq 10) {
        $segLen = $i - $start
        if ($segLen -gt 0) {
          $seg = New-Object byte[] $segLen
          [Array]::Copy($all, $start, $seg, 0, $segLen)
          $line = Decode-Smart $seg
          $line = $line.TrimStart([char]0xFEFF).TrimEnd([char]13)
          if ($line -ne '') { Add-ParsedLine $line }
        }
        $start = $i + 1
      }
    }
    # 剩余未完成行暂存字节
    if ($start -lt $all.Length) {
      $rest = New-Object byte[] ($all.Length - $start)
      [Array]::Copy($all, $start, $rest, 0, $rest.Length)
      $script:pendingBytes = $rest
    } else {
      $script:pendingBytes = New-Object byte[] 0
    }
    if ($FlushPending -and $script:pendingBytes.Length -gt 0) {
      $line = Decode-Smart $script:pendingBytes
      $line = $line.TrimStart([char]0xFEFF).TrimEnd([char]13)
      if ($line -ne '') { Add-ParsedLine $line }
      $script:pendingBytes = New-Object byte[] 0
    }
    if ($script:logLines.Count -gt 3000) {
      $script:logLines.RemoveRange(0, 500)
      if ($script:renderedCount -gt 500) { $script:renderedCount -= 500 } else { $script:renderedCount = 0 }
    }
  } catch {}
}
function Render-NewLogs {
  for ($i = $script:renderedCount; $i -lt $script:logLines.Count; $i++) {
    Add-LogLine $script:logLines[$i]
  }
  $script:renderedCount = $script:logLines.Count
}

function Add-LogLine([string]$line) {
  $rtb = $script:richLog
  $rtb.SelectionStart = $rtb.TextLength
  $rtb.SelectionLength = 0
  $c = '#CBD5E1'
  if ($line -match '\[错误\]|\[FAIL\]|失败（退出码') { $c = '#F87171' }
  elseif ($line -match '\[PASS\]|安装成功|成功|✓') { $c = '#4ADE80' }
  elseif ($line -match '\[注意\]|未通过|已取消|跳过') { $c = '#FBBF24' }
  elseif ($line -match '\[任务\]|\[MCP\]|\[中心\]|====') { $c = '#60A5FA' }
  elseif ($line -match '^\[[0-9]+/[0-9]+\]') { $c = '#60A5FA' }
  $rtb.SelectionColor = [System.Drawing.ColorTranslator]::FromHtml($c)
  $rtb.AppendText($line + $script:nl)
  $rtb.SelectionStart = $rtb.TextLength
  $rtb.ScrollToCaret()
}

function Start-McpWindow {
  $node = Get-Command node -ErrorAction SilentlyContinue
  if (-not $node) { return '未检测到 Node.js，MCP 需要 Node.js 24+' }
  $target = Join-Path $root 'server\mcp\index.js'
  if (-not (Test-Path -LiteralPath $target)) { return '未找到 server\mcp\index.js' }
  Start-Process -FilePath 'cmd.exe' -ArgumentList @('/k', 'chcp 65001>nul & cd /d "' + $root + '" & echo [MCP] 正在运行（stdio），关闭本窗口即停止 & node "' + $target + '"')
  Add-LogLine '[MCP] 已在独立窗口启动'
  return $null
}

function New-DesktopShortcut {
  try {
    $ws = New-Object -ComObject WScript.Shell
    $desktop = [Environment]::GetFolderPath('Desktop')
    $lnkPath = Join-Path $desktop '模板助手安装向导.lnk'
    $lnk = $ws.CreateShortcut($lnkPath)
    $exe = Join-Path $root '安装中心.exe'
    $ico = Join-Path $root '安装中心.ico'
    if (Test-Path -LiteralPath $exe) { $lnk.TargetPath = $exe } else { $lnk.TargetPath = Join-Path $root '安装中心.bat' }
    $lnk.WorkingDirectory = $root
    $lnk.Description = '模板助手 · 安装向导'
    if (Test-Path -LiteralPath $ico) { $lnk.IconLocation = "$ico,0" }
    $lnk.Save()
    return $null
  } catch { return $_.Exception.Message }
}

function Open-Target([string]$target) {
  switch ($target) {
    'install_dir' {
      $d = Join-Path $env:LOCALAPPDATA 'PPT-AI-Addin'
      if (-not (Test-Path -LiteralPath $d)) { $null = New-Item -ItemType Directory -Path $d -Force }
      [System.Diagnostics.Process]::Start('explorer.exe', ('"' + $d + '"')) | Out-Null
    }
    'template_lib' {
      $d = Join-Path ([Environment]::GetFolderPath('MyDocuments')) 'PPT模板库'
      if (-not (Test-Path -LiteralPath $d)) { $null = New-Item -ItemType Directory -Path $d -Force }
      [System.Diagnostics.Process]::Start('explorer.exe', ('"' + $d + '"')) | Out-Null
    }
    'doc' { [System.Diagnostics.Process]::Start((Join-Path $root '安装说明.md')) | Out-Null }
    'nodejs' { [System.Diagnostics.Process]::Start('https://nodejs.org/zh-cn/download') | Out-Null }
    'powerpoint' { [System.Diagnostics.Process]::Start('powerpnt.exe') | Out-Null }
  }
}

function Cancel-Job {
  if ($script:job -and -not $script:job.Proc.HasExited) {
    $script:cancelled = $true
    & taskkill /PID $script:job.Proc.Id /T /F 2>$null | Out-Null
    Add-LogLine '[任务] 已发送取消指令（正在终止进程树）…'
  }
  if ($script:lastTmpBat -and (Test-Path -LiteralPath $script:lastTmpBat)) { Remove-Item -LiteralPath $script:lastTmpBat -Force -ErrorAction SilentlyContinue }
  $script:lastTmpBat = $null
}

# ============================================================
#  自检模式（不弹窗口，用于验证脚本）
# ============================================================
if ($SelfTest) {
  $env = Get-EnvCheck
  Write-Output ('ENV: ' + ($env | ConvertTo-Json -Compress -Depth 3))
  $tmp = Join-Path $env:TEMP ('ppt-addin-ui-selftest-' + [guid]::NewGuid().ToString('N') + '.bat')
  $content = '@chcp 65001 >nul' + $script:crlf + '@echo off' + $script:crlf + 'echo [0/6] 检查运行环境' + $script:crlf + 'echo [2/6] 启动本地服务' + $script:crlf + 'echo [PASS] 环境正常' + $script:crlf + 'exit /b 0'
  [System.IO.File]::WriteAllText($tmp, $content, (New-Object System.Text.UTF8Encoding($false)))
  $ok = Start-BatJob 'selftest' $tmp ''
  Start-Sleep -Milliseconds 1500
  Sync-JobLog
  $j = $script:job
  Write-Output ('JOB-START: ' + $ok)
  Write-Output ('JOB-STATE: step=' + $j.Step + ' total=' + $j.TotalSteps + ' exited=' + $j.Proc.HasExited + ' exitCode=' + $j.Proc.ExitCode)
  Write-Output ('LOG-LINES: ' + $script:logLines.Count)
  $hasMoji = $false
  foreach ($l in $script:logLines) { if ($l.IndexOf([char]0xFFFD) -ge 0) { $hasMoji = $true } }
  Write-Output ('NO-MOJIBAKE: ' + (-not $hasMoji))
  Write-Output ('STAGE-LABEL: ' + $j.StepLabel)
  Write-Output ('STAGE-CLEAN: ' + ($j.StepLabel -eq '启动本地服务'))
  Remove-Item -LiteralPath $tmp -Force -ErrorAction SilentlyContinue
  exit 0
}

# ============================================================
#  构建向导窗口
# ============================================================
# ---- 视觉基调（简洁中性色板）----
$cAccent  = [System.Drawing.ColorTranslator]::FromHtml('#2563EB')
$cText    = [System.Drawing.ColorTranslator]::FromHtml('#111827')
$cText2   = [System.Drawing.ColorTranslator]::FromHtml('#6B7280')
$cText3   = [System.Drawing.ColorTranslator]::FromHtml('#9CA3AF')
$cLine    = [System.Drawing.ColorTranslator]::FromHtml('#E5E7EB')
$cOk      = [System.Drawing.ColorTranslator]::FromHtml('#16A34A')
$cWarn    = [System.Drawing.ColorTranslator]::FromHtml('#D97706')
$cErr     = [System.Drawing.ColorTranslator]::FromHtml('#DC2626')
$cSoft    = [System.Drawing.ColorTranslator]::FromHtml('#F8FAFC')
$cAccentSoft = [System.Drawing.ColorTranslator]::FromHtml('#EEF4FF')
$cLogBg   = [System.Drawing.ColorTranslator]::FromHtml('#0F172A')
$cLogFg   = [System.Drawing.ColorTranslator]::FromHtml('#CBD5E1')

$font   = New-Object System.Drawing.Font('Microsoft YaHei UI', 9.5)
$fontB  = New-Object System.Drawing.Font('Microsoft YaHei UI', 9.5, [System.Drawing.FontStyle]::Bold)
$fontT  = New-Object System.Drawing.Font('Microsoft YaHei UI', 18, [System.Drawing.FontStyle]::Bold)
$fontH  = New-Object System.Drawing.Font('Microsoft YaHei UI', 12, [System.Drawing.FontStyle]::Bold)
$fontGlyph = New-Object System.Drawing.Font('Segoe UI Symbol', 10)
$fontBigGlyph = New-Object System.Drawing.Font('Segoe UI Symbol', 26, [System.Drawing.FontStyle]::Bold)
$fontMono = New-Object System.Drawing.Font('Consolas', 9.5)

function New-Button([string]$text, [int]$x, [int]$y, [int]$w, [int]$h, [string]$style) {
  $b = New-Object System.Windows.Forms.Button
  $b.SetBounds($x, $y, $w, $h)
  $b.Text = $text
  $b.FlatStyle = 'Flat'
  $b.Font = $font
  $b.Cursor = [System.Windows.Forms.Cursors]::Hand
  if ($style -eq 'primary') {
    $b.BackColor = $cAccent
    $b.ForeColor = [System.Drawing.Color]::White
    $b.FlatAppearance.BorderSize = 0
  } else {
    $b.BackColor = [System.Drawing.Color]::White
    $b.ForeColor = $cText
    $b.FlatAppearance.BorderColor = [System.Drawing.ColorTranslator]::FromHtml('#D1D5DB')
    $b.FlatAppearance.BorderSize = 1
  }
  return $b
}

function New-Row([int]$y, [string]$icon, [string]$iconColor, [string]$name, [string]$sub) {
  $p = New-Object System.Windows.Forms.Panel
  $p.SetBounds(40, $y, 640, 52)
  $p.BackColor = [System.Drawing.Color]::White
  $ic = New-Object System.Windows.Forms.Label
  $ic.SetBounds(2, 14, 22, 22)
  $ic.Text = $icon
  $ic.TextAlign = 'MiddleCenter'
  $ic.Font = $fontGlyph
  $ic.ForeColor = [System.Drawing.ColorTranslator]::FromHtml($iconColor)
  $nm = New-Object System.Windows.Forms.Label
  $nm.SetBounds(34, 15, 180, 22)
  $nm.Text = $name
  $nm.Font = $fontB
  $nm.ForeColor = $cText
  $sb = New-Object System.Windows.Forms.Label
  $sb.SetBounds(222, 17, 350, 20)
  $sb.Text = $sub
  $sb.ForeColor = $cText2
  $sb.TextAlign = 'MiddleLeft'
  $p.Controls.Add($ic); $p.Controls.Add($nm); $p.Controls.Add($sb)
  return $p
}

$form = New-Object System.Windows.Forms.Form
$form.Text = '模板助手 · 安装向导'
$form.ClientSize = New-Object System.Drawing.Size(940, 640)
$form.FormBorderStyle = 'FixedSingle'
$form.MaximizeBox = $false
$form.StartPosition = 'CenterScreen'
$form.BackColor = [System.Drawing.Color]::White
$form.Font = $font
$form.KeyPreview = $true
if ($script:appIcon) { $form.Icon = $script:appIcon; $form.ShowIcon = $true }

# ---- 顶部标题栏（白底 + 细分割线） ----
$hdr = New-Object System.Windows.Forms.Panel
$hdr.SetBounds(0, 0, 940, 72)
$hdr.BackColor = [System.Drawing.Color]::White
if ($script:appIcon) {
  $logo = New-Object System.Windows.Forms.PictureBox
  $logo.SetBounds(18, 16, 40, 40)
  $logo.SizeMode = 'StretchImage'
  $logo.Image = $script:appIcon.ToBitmap()
} else {
  $logo = New-Object System.Windows.Forms.Label
  $logo.SetBounds(18, 16, 40, 40)
  $logo.Text = 'P'
  $logo.Font = New-Object System.Drawing.Font('Microsoft YaHei UI', 16, [System.Drawing.FontStyle]::Bold)
  $logo.ForeColor = [System.Drawing.Color]::White
  $logo.TextAlign = 'MiddleCenter'
  $logo.BackColor = $cAccent
}
$lblHdrTitle = New-Object System.Windows.Forms.Label
$lblHdrTitle.SetBounds(72, 14, 600, 26)
$lblHdrTitle.Text = '模板助手 · 安装向导'
$lblHdrTitle.Font = $fontH
$lblHdrTitle.ForeColor = $cText
$lblHdrSub = New-Object System.Windows.Forms.Label
$lblHdrSub.SetBounds(72, 43, 700, 18)
$lblHdrSub.Text = '环境检测 · 安装 · 自检'
$lblHdrSub.ForeColor = $cText2
$hdrLine = New-Object System.Windows.Forms.Panel
$hdrLine.SetBounds(0, 71, 940, 1)
$hdrLine.BackColor = $cLine
$hdr.Controls.Add($logo); $hdr.Controls.Add($lblHdrTitle); $hdr.Controls.Add($lblHdrSub); $hdr.Controls.Add($hdrLine)

# ---- 左侧步骤栏（状态点 + 细连接线，简洁） ----
$rail = New-Object System.Windows.Forms.Panel
$rail.SetBounds(0, 72, 216, 510)
$rail.BackColor = [System.Drawing.ColorTranslator]::FromHtml('#FAFAFB')
$script:railSteps = @()
$stepDefs = @(
  @{ t = '欢迎' }, @{ t = '环境检查' }, @{ t = '安装选项' }, @{ t = '正在安装' }, @{ t = '完成' }
)
$conn = New-Object System.Windows.Forms.Panel
$conn.SetBounds(25, 40, 1, 224)
$conn.BackColor = $cLine
$rail.Controls.Add($conn)
for ($i = 0; $i -lt 5; $i++) {
  $y = 18 + $i * 56
  $g = New-Object System.Windows.Forms.Label
  $g.SetBounds(16, $y + 12, 20, 20)
  $g.Text = '○'
  $g.Font = $fontGlyph
  $g.TextAlign = 'MiddleCenter'
  $g.ForeColor = [System.Drawing.ColorTranslator]::FromHtml('#C9D2DE')
  $t = New-Object System.Windows.Forms.Label
  $t.SetBounds(44, $y + 13, 150, 20)
  $t.Text = $stepDefs[$i].t
  $t.Font = $font
  $t.ForeColor = $cText3
  $rail.Controls.Add($g); $rail.Controls.Add($t)
  $script:railSteps += ,@($g, $t)
}
$railLine = New-Object System.Windows.Forms.Panel
$railLine.SetBounds(215, 0, 1, 510)
$railLine.BackColor = $cLine
$rail.Controls.Add($railLine)

# ---- 内容区 ----
$main = New-Object System.Windows.Forms.Panel
$main.SetBounds(216, 72, 724, 510)
$main.BackColor = [System.Drawing.Color]::White

function New-StepPanel {
  $p = New-Object System.Windows.Forms.Panel
  $p.SetBounds(0, 0, 724, 510)
  $p.BackColor = [System.Drawing.Color]::White
  $p.Visible = $false
  return $p
}

# ---- 步骤 1：欢迎 ----
$pnlWelcome = New-StepPanel
$w1 = New-Object System.Windows.Forms.Label
$w1.SetBounds(40, 30, 640, 34)
$w1.Text = '欢迎使用「模板助手」'
$w1.Font = $fontT
$w1.ForeColor = $cText
$w2 = New-Object System.Windows.Forms.Label
$w2.SetBounds(40, 74, 640, 22)
$w2.Text = '本向导将依次完成环境检测、构建、安装、注册与自检。'
$w2.ForeColor = $cText2
$flow = New-Object System.Windows.Forms.Panel
$flow.SetBounds(40, 116, 640, 172)
$flow.BackColor = [System.Drawing.Color]::White
$items = @(
  '检查运行环境（Node.js、构建产物、已安装状态）',
  '构建并打包发布产物（或使用现有产物）',
  '安装到 %LOCALAPPDATA%PPT-AI-Addin',
  '注册到 PowerPoint 并设置登录自启动',
  '启动本地服务并完成安装自检'
)
for ($i = 0; $i -lt $items.Count; $i++) {
  $dot = New-Object System.Windows.Forms.Label
  $dot.SetBounds(4, 6 + $i * 32, 18, 18)
  $dot.Text = '•'
  $dot.Font = $fontB
  $dot.TextAlign = 'MiddleCenter'
  $dot.ForeColor = $cText3
  $li = New-Object System.Windows.Forms.Label
  $li.SetBounds(28, 4 + $i * 32, 600, 22)
  $li.Text = $items[$i]
  $li.ForeColor = $cText
  $flow.Controls.Add($dot); $flow.Controls.Add($li)
}
$w3 = New-Object System.Windows.Forms.Label
$w3.SetBounds(40, 308, 640, 48)
$w3.Text = '需要 Windows 10 / 11 与 Office 2016 及以上。' + $script:nl + '全新构建需要 Node.js；选择「使用现有构建产物」则不需要。'
$w3.ForeColor = $cText2
$w3.BackColor = $cSoft
$w3.Padding = New-Object System.Windows.Forms.Padding(12, 6, 12, 6)
$btnUninstallLink = New-Object System.Windows.Forms.Button
$btnUninstallLink.SetBounds(40, 380, 200, 28)
$btnUninstallLink.Text = '卸载模板助手'
$btnUninstallLink.FlatStyle = 'Flat'
$btnUninstallLink.FlatAppearance.BorderSize = 0
$btnUninstallLink.TextAlign = 'MiddleLeft'
$btnUninstallLink.ForeColor = $cErr
$btnUninstallLink.Cursor = [System.Windows.Forms.Cursors]::Hand
$pnlWelcome.Controls.Add($w1); $pnlWelcome.Controls.Add($w2); $pnlWelcome.Controls.Add($flow)
$pnlWelcome.Controls.Add($w3); $pnlWelcome.Controls.Add($btnUninstallLink)

# ---- 步骤 2：环境检查 ----
$pnlCheck = New-StepPanel
$c1 = New-Object System.Windows.Forms.Label
$c1.SetBounds(40, 24, 640, 30)
$c1.Text = '环境检查'
$c1.Font = $fontT
$c1.ForeColor = $cText
$c2 = New-Object System.Windows.Forms.Label
$c2.SetBounds(40, 62, 420, 20)
$c2.Text = '检测本机环境与已安装状态：'
$c2.ForeColor = $cText2
$script:rowNode   = New-Row 104 '…' '#9CA3AF' 'Node.js' '检测中…'
$script:rowRel    = New-Row 158 '…' '#9CA3AF' '发布产物（release.json）' '检测中…'
$script:rowInst   = New-Row 212 '…' '#9CA3AF' '已安装版本' '检测中…'
$script:rowReg    = New-Row 266 '…' '#9CA3AF' 'PowerPoint 注册' '检测中…'
$script:rowSvc    = New-Row 320 '…' '#9CA3AF' '本地服务' '检测中…'
$script:lblEnvNote = New-Object System.Windows.Forms.Label
$script:lblEnvNote.SetBounds(40, 388, 640, 56)
$script:lblEnvNote.Visible = $false
$script:lblEnvNote.Padding = New-Object System.Windows.Forms.Padding(12, 6, 12, 6)
$btnDlNode = New-Object System.Windows.Forms.Button
$btnDlNode.SetBounds(552, 116, 128, 28)
$btnDlNode.Text = '下载 Node.js'
$btnDlNode.Visible = $false
$btnDlNode.FlatStyle = 'Flat'
$btnDlNode.ForeColor = $cAccent
$btnDlNode.FlatAppearance.BorderColor = [System.Drawing.ColorTranslator]::FromHtml('#BFDBFE')
$btnDlNode.FlatAppearance.BorderSize = 1
$btnRecheck = New-Object System.Windows.Forms.Button
$btnRecheck.SetBounds(560, 60, 120, 24)
$btnRecheck.Text = '重新检测'
$btnRecheck.FlatStyle = 'Flat'
$btnRecheck.FlatAppearance.BorderSize = 0
$btnRecheck.TextAlign = 'MiddleRight'
$btnRecheck.ForeColor = $cAccent
$btnRecheck.Cursor = [System.Windows.Forms.Cursors]::Hand
$pnlCheck.Controls.Add($c1); $pnlCheck.Controls.Add($c2)
$pnlCheck.Controls.Add($script:rowNode); $pnlCheck.Controls.Add($script:rowRel)
$pnlCheck.Controls.Add($script:rowInst); $pnlCheck.Controls.Add($script:rowReg)
$pnlCheck.Controls.Add($script:rowSvc); $pnlCheck.Controls.Add($script:lblEnvNote)
$pnlCheck.Controls.Add($btnDlNode); $pnlCheck.Controls.Add($btnRecheck)

# ---- 步骤 3：安装选项 ----
$pnlOptions = New-StepPanel
$o1 = New-Object System.Windows.Forms.Label
$o1.SetBounds(40, 24, 640, 30)
$o1.Text = '安装选项'
$o1.Font = $fontT
$o1.ForeColor = $cText
$o2 = New-Object System.Windows.Forms.Label
$o2.SetBounds(40, 62, 640, 20)
$o2.Text = '选择安装方式与附加选项：'
$o2.ForeColor = $cText2
function New-OptionCard([int]$y, [string]$title, [string]$sub) {
  $card = New-Object System.Windows.Forms.Panel
  $card.SetBounds(40, $y, 640, 72)
  $card.BackColor = [System.Drawing.Color]::White
  $card.BorderStyle = 'FixedSingle'
  $bar = New-Object System.Windows.Forms.Panel
  $bar.SetBounds(0, 0, 4, 72)
  $bar.BackColor = $cAccent
  $bar.Visible = $false
  $rad = New-Object System.Windows.Forms.RadioButton
  $rad.SetBounds(22, 26, 20, 20)
  $txt = New-Object System.Windows.Forms.Label
  $txt.SetBounds(54, 12, 560, 22)
  $txt.Text = $title
  $txt.Font = $fontB
  $txt.ForeColor = $cText
  $subL = New-Object System.Windows.Forms.Label
  $subL.SetBounds(54, 40, 560, 20)
  $subL.Text = $sub
  $subL.ForeColor = $cText2
  $card.Controls.Add($bar); $card.Controls.Add($rad)
  $card.Controls.Add($txt); $card.Controls.Add($subL)
  $card.Add_Click({ $rad.Checked = $true })
  $txt.Add_Click({ $rad.Checked = $true })
  $subL.Add_Click({ $rad.Checked = $true })
  return ,@($card, $rad, $bar)
}
$cFull = New-OptionCard 100 '全新构建并安装（推荐）' '完整构建：版本号 → 前端构建 → exe 打包 → 安装 → 注册 → 自检（需 Node.js）。'
$script:cardFull = $cFull[0]; $script:radFull = $cFull[1]; $script:cardFullBar = $cFull[2]
$cSkip = New-OptionCard 184 '使用现有构建产物（跳过构建）' '不运行构建，直接安装当前目录已有的构建产物。'
$script:cardSkip = $cSkip[0]; $script:radSkip = $cSkip[1]; $script:cardSkipBar = $cSkip[2]
$script:radFull.Checked = $true
$script:chkTests = New-Object System.Windows.Forms.CheckBox
$script:chkTests.SetBounds(40, 280, 400, 22)
$script:chkTests.Text = '包含 npm 测试（较慢）'
$script:chkTests.ForeColor = $cText
$script:chkSkipE2e = New-Object System.Windows.Forms.CheckBox
$script:chkSkipE2e.SetBounds(40, 310, 400, 22)
$script:chkSkipE2e.Text = '跳过 E2E 自检（约 10 秒）'
$script:chkSkipE2e.ForeColor = $cText
$script:lblOptNote = New-Object System.Windows.Forms.Label
$script:lblOptNote.SetBounds(40, 344, 640, 48)
$script:lblOptNote.Visible = $false
$script:lblOptNote.Padding = New-Object System.Windows.Forms.Padding(12, 6, 12, 6)
$script:lblOptSum = New-Object System.Windows.Forms.Label
$script:lblOptSum.SetBounds(40, 400, 640, 52)
$script:lblOptSum.BackColor = $cSoft
$script:lblOptSum.ForeColor = $cText2
$script:lblOptSum.Padding = New-Object System.Windows.Forms.Padding(12, 0, 12, 0)
$pnlOptions.Controls.Add($o1); $pnlOptions.Controls.Add($o2)
$pnlOptions.Controls.Add($script:cardFull); $pnlOptions.Controls.Add($script:cardSkip)
$pnlOptions.Controls.Add($script:chkTests); $pnlOptions.Controls.Add($script:chkSkipE2e)
$pnlOptions.Controls.Add($script:lblOptNote); $pnlOptions.Controls.Add($script:lblOptSum)

# ---- 步骤 4：正在安装 ----
$pnlInstall = New-StepPanel
$i1 = New-Object System.Windows.Forms.Label
$i1.SetBounds(40, 24, 640, 30)
$i1.Text = '正在安装'
$i1.Font = $fontT
$i1.ForeColor = $cText
$script:lblInstLead = New-Object System.Windows.Forms.Label
$script:lblInstLead.SetBounds(40, 62, 640, 20)
$script:lblInstLead.Text = '正在执行安装任务，请稍候…'
$script:lblInstLead.ForeColor = $cText2
$script:lblStage = New-Object System.Windows.Forms.Label
$script:lblStage.SetBounds(40, 98, 420, 22)
$script:lblStage.Font = $fontB
$script:lblStage.ForeColor = $cText
$script:lblElapsed = New-Object System.Windows.Forms.Label
$script:lblElapsed.SetBounds(500, 100, 180, 20)
$script:lblElapsed.TextAlign = 'MiddleRight'
$script:lblElapsed.ForeColor = $cText2
$script:progressBar = New-Object System.Windows.Forms.ProgressBar
$script:progressBar.SetBounds(40, 130, 640, 14)
$script:progressBar.Minimum = 0
$script:progressBar.Maximum = 100
$script:progressBar.Style = 'Continuous'
$script:richLog = New-Object System.Windows.Forms.RichTextBox
$script:richLog.SetBounds(40, 160, 640, 330)
$script:richLog.ReadOnly = $true
$script:richLog.BackColor = $cLogBg
$script:richLog.ForeColor = $cLogFg
$script:richLog.BorderStyle = 'FixedSingle'
$script:richLog.Font = $fontMono
$script:richLog.WordWrap = $true
$pnlInstall.Controls.Add($i1); $pnlInstall.Controls.Add($script:lblInstLead)
$pnlInstall.Controls.Add($script:lblStage); $pnlInstall.Controls.Add($script:lblElapsed)
$pnlInstall.Controls.Add($script:progressBar); $pnlInstall.Controls.Add($script:richLog)

# ---- 步骤 5：完成 ----
$pnlDone = New-StepPanel
$lblDoneCheck = New-Object System.Windows.Forms.Label
$lblDoneCheck.SetBounds(328, 30, 64, 64)
$lblDoneCheck.Text = '✓'
$lblDoneCheck.Font = $fontBigGlyph
$lblDoneCheck.TextAlign = 'MiddleCenter'
$lblDoneCheck.ForeColor = $cOk
$script:lblDoneTitle = New-Object System.Windows.Forms.Label
$script:lblDoneTitle.SetBounds(40, 106, 640, 38)
$script:lblDoneTitle.TextAlign = 'MiddleCenter'
$script:lblDoneTitle.Font = $fontT
$script:lblDoneSub = New-Object System.Windows.Forms.Label
$script:lblDoneSub.SetBounds(40, 152, 640, 52)
$script:lblDoneSub.TextAlign = 'MiddleCenter'
$script:lblDoneSub.ForeColor = $cText2
$doneCard = New-Object System.Windows.Forms.Panel
$doneCard.SetBounds(40, 220, 640, 128)
$doneCard.BackColor = $cSoft
$script:lblDoneList = New-Object System.Windows.Forms.Label
$script:lblDoneList.SetBounds(56, 10, 608, 106)
$script:lblDoneList.ForeColor = [System.Drawing.ColorTranslator]::FromHtml('#374151')
$doneCard.Controls.Add($script:lblDoneList)
$script:btnPpt = New-Button '打开 PowerPoint' 103 372 130 36 'primary'
$script:btnMcp = New-Button '启动 MCP' 241 372 110 36 'outline'
$script:btnDir = New-Button '打开安装目录' 359 372 120 36 'outline'
$script:btnShortcut = New-Button '创建桌面快捷方式' 487 372 130 36 'outline'
$script:btnViewLog = New-Button '查看日志' 240 372 120 36 'outline'
$script:btnRetry = New-Button '返回重试' 370 372 120 36 'outline'
$pnlDone.Controls.Add($lblDoneCheck); $pnlDone.Controls.Add($script:lblDoneTitle)
$pnlDone.Controls.Add($script:lblDoneSub); $pnlDone.Controls.Add($doneCard)
$pnlDone.Controls.Add($script:btnPpt); $pnlDone.Controls.Add($script:btnMcp)
$pnlDone.Controls.Add($script:btnDir); $pnlDone.Controls.Add($script:btnShortcut)
$pnlDone.Controls.Add($script:btnViewLog); $pnlDone.Controls.Add($script:btnRetry)

$main.Controls.Add($pnlWelcome); $main.Controls.Add($pnlCheck)
$main.Controls.Add($pnlOptions); $main.Controls.Add($pnlInstall); $main.Controls.Add($pnlDone)

# ---- 底部按钮 ----
$foot = New-Object System.Windows.Forms.Panel
$foot.SetBounds(0, 582, 940, 58)
$foot.BackColor = [System.Drawing.Color]::White
$footLine = New-Object System.Windows.Forms.Panel
$footLine.SetBounds(0, 0, 940, 1)
$footLine.BackColor = $cLine
$script:lblStepPos = New-Object System.Windows.Forms.Label
$script:lblStepPos.SetBounds(24, 19, 160, 20)
$script:lblStepPos.Text = '步骤 1 / 5'
$script:lblStepPos.ForeColor = $cText3
$script:btnBack = New-Button '上一步' 552 13 96 32 'outline'
$script:btnCancel = New-Button '取消' 656 13 96 32 'outline'
$script:btnNext = New-Button '下一步' 764 11 128 36 'primary'
$foot.Controls.Add($footLine); $foot.Controls.Add($script:lblStepPos)
$foot.Controls.Add($script:btnBack); $foot.Controls.Add($script:btnCancel); $foot.Controls.Add($script:btnNext)

$form.AcceptButton = $script:btnNext
$form.CancelButton = $script:btnCancel

$form.Controls.Add($hdr); $form.Controls.Add($rail); $form.Controls.Add($main); $form.Controls.Add($foot)

# ============================================================
#  逻辑
# ============================================================
function Set-Step([int]$n) {
  $script:step = $n
  $panels = @($pnlWelcome, $pnlCheck, $pnlOptions, $pnlInstall, $pnlDone)
  for ($i = 0; $i -lt 5; $i++) { $panels[$i].Visible = ($i -eq $n) }
  for ($i = 0; $i -lt 5; $i++) {
    $r = $script:railSteps[$i]
    if ($i -eq $n) {
      $r[0].Text = '●'; $r[0].ForeColor = $cAccent
      $r[1].Font = $fontB; $r[1].ForeColor = $cText
    } elseif ($i -lt $n) {
      $r[0].Text = '✓'; $r[0].ForeColor = $cOk
      $r[1].Font = $font; $r[1].ForeColor = $cText2
    } else {
      $r[0].Text = '○'; $r[0].ForeColor = [System.Drawing.ColorTranslator]::FromHtml('#C9D2DE')
      $r[1].Font = $font; $r[1].ForeColor = $cText3
    }
  }
  $script:lblStepPos.Text = '步骤 ' + ($n + 1) + ' / 5'
  if ($n -eq 3) {
    $script:btnBack.Visible = $false
    $script:btnNext.Visible = $false
    $script:btnCancel.Visible = $true
    $running = ($script:job -and -not $script:job.Proc.HasExited)
    $script:btnCancel.Text = if ($running) { '取消任务' } else { '退出' }
    $script:btnCancel.ForeColor = if ($running) { $cErr } else { $cText }
  } elseif ($n -eq 4) {
    $script:btnBack.Visible = $false
    $script:btnCancel.Visible = $false
    $script:btnNext.Visible = $true
    $script:btnNext.Text = '完成'
  } else {
    $script:btnBack.Visible = ($n -ge 1)
    $script:btnCancel.Visible = $true
    $script:btnCancel.Text = '取消'
    $script:btnCancel.ForeColor = $cText
    $script:btnNext.Visible = $true
    if ($n -eq 2) { $script:btnNext.Text = '开始安装' } else { $script:btnNext.Text = '下一步' }
    $script:btnNext.Enabled = if ($n -eq 1) { $script:envOK } else { $true }
  }
  if ($n -eq 1) { Render-Check }
  if ($n -eq 2) { Enter-Options }
  if ($n -eq 3) { Enter-Install }
  if ($n -eq 4) { Render-Done }
}

function Render-Check {
  $env = Get-EnvCheck
  $script:env = $env
  function Row-Set($row, [string]$icon, [string]$color, [string]$sub) {
    $row.Controls[0].Text = $icon
    $row.Controls[0].ForeColor = [System.Drawing.ColorTranslator]::FromHtml($color)
    $row.Controls[2].Text = $sub
  }
  if ($env.NodePresent) {
    Row-Set $script:rowNode '✓' '#16A34A' (('v' + $env.NodeVersion.TrimStart('v')) + $(if ($env.NpmVersion) { '   ·   npm ' + $env.NpmVersion } else { '' }))
    $btnDlNode.Visible = $false
  } else {
    Row-Set $script:rowNode '✗' '#DC2626' '未检测到 Node.js — 无法进行全新构建'
    $btnDlNode.Visible = $true
  }
  if ($env.RelPresent) {
    Row-Set $script:rowRel '✓' '#16A34A' ('v' + $env.RelVersion + $(if ($env.RelBuilt) { '   ·   构建于 ' + $env.RelBuilt.Replace('T', ' ').Substring(0, 16) } else { '' }))
  } else {
    Row-Set $script:rowRel '○' '#D97706' '未找到 release.json — 需要全新构建（需 Node.js）'
  }
  if ($env.InstPresent) {
    $same = $env.RelPresent -and $env.RelVersion -eq $env.InstVersion
    Row-Set $script:rowInst '✓' '#16A34A' ('v' + $env.InstVersion + $(if ($same) { '   ·   与发布版一致' } else { '   ·   与发布版不同，建议重装' }))
  } else {
    Row-Set $script:rowInst '○' '#D97706' '未安装（本次为全新安装）'
  }
  if ($env.RegOK) {
    Row-Set $script:rowReg '✓' '#16A34A' '已注册（当前用户 WEF）'
  } else {
    Row-Set $script:rowReg '○' '#D97706' '未注册 — 安装后自动注册'
  }
  if ($env.SvcVersion) {
    Row-Set $script:rowSvc '✓' '#16A34A' ('运行中 · v' + $env.SvcVersion + ' · 端口 ' + $env.SvcPort)
  } elseif ($env.SvcProcess) {
    Row-Set $script:rowSvc '○' '#D97706' '进程在，API 未通 — 安装后会自动修复'
  } else {
    Row-Set $script:rowSvc '○' '#D97706' '未运行 — 安装后自动启动'
  }
  $script:envOK = ($env.NodePresent -or $env.RelPresent)
  $note = $script:lblEnvNote
  if (-not $script:envOK) {
    $note.Visible = $true
    $note.BackColor = [System.Drawing.ColorTranslator]::FromHtml('#FEF2F2')
    $note.ForeColor = [System.Drawing.ColorTranslator]::FromHtml('#B91C1C')
    $note.Text = '无法继续：既没有 Node.js，也没有构建产物。' + $script:nl + '请点击「下载 Node.js」安装，或将完整构建产物（dist-exe、addin\dist、manifest.xml、release.json）放入本项目目录后重新检测。'
  } elseif (-not $env.NodePresent) {
    $note.Visible = $true
    $note.BackColor = [System.Drawing.ColorTranslator]::FromHtml('#FFFBEB')
    $note.ForeColor = [System.Drawing.ColorTranslator]::FromHtml('#92400E')
    $note.Text = 'Node.js 未安装：将自动使用「现有构建产物」方式（跳过构建）安装。'
  } elseif (-not $env.RelPresent) {
    $note.Visible = $true
    $note.BackColor = [System.Drawing.ColorTranslator]::FromHtml('#FFFBEB')
    $note.ForeColor = [System.Drawing.ColorTranslator]::FromHtml('#92400E')
    $note.Text = '未找到构建产物：将自动进行全新构建（需要联网下载依赖，首次较慢）。'
  } else {
    $note.Visible = $false
  }
  $script:btnNext.Enabled = $script:envOK
}

function Enter-Options {
  $env = $script:env
  $nodeOK = $env.NodePresent
  $relOK  = $env.RelPresent
  $script:radFull.Enabled = $nodeOK
  $script:radSkip.Enabled = $relOK
  if (-not $nodeOK -and $relOK) { $script:radSkip.Checked = $true; $script:radFull.Checked = $false }
  elseif ($nodeOK -and -not $relOK) { $script:radFull.Checked = $true; $script:radSkip.Checked = $false }
  Sync-Options
}

function Update-OptionCards {
  $full = $script:radFull.Checked
  $script:cardFull.BackColor = if ($full) { $cAccentSoft } else { [System.Drawing.Color]::White }
  $script:cardFullBar.Visible = $full
  $script:cardSkip.BackColor = if (-not $full) { $cAccentSoft } else { [System.Drawing.Color]::White }
  $script:cardSkipBar.Visible = (-not $full)
}

function Sync-Options {
  $skip = $script:radSkip.Checked
  $script:chkTests.Enabled = -not $skip
  if ($skip) { $script:chkTests.Checked = $false }
  Update-OptionCards
  $note = $script:lblOptNote
  if ($skip) {
    $note.Visible = $true
    $note.BackColor = [System.Drawing.ColorTranslator]::FromHtml('#FFFBEB')
    $note.ForeColor = [System.Drawing.ColorTranslator]::FromHtml('#92400E')
    $note.Text = '已选择「使用现有构建产物」：不运行构建，直接安装当前目录已有产物。'
  } else {
    $note.Visible = $false
  }
  if ($skip) {
    $script:lblOptSum.Text = '将执行：安装现有产物 → 注册 PowerPoint → 设置自启动 → 启动服务 → 自检。'
  } else {
    $script:lblOptSum.Text = '将执行：完整构建（版本号 + 前端构建 + exe 打包）' + $(if ($script:chkTests.Checked) { '（含 npm 测试）' } else { '（不含 npm 测试）' }) + $(if ($script:chkSkipE2e.Checked) { '，跳过 E2E 自检' } else { '，含 E2E 自检约 10 秒' }) + ' → 安装 → 注册 → 自检。'
  }
}

function Enter-Install {
  $script:richLog.Clear()
  $script:renderedCount = 0
  if ($script:mode -eq 'uninstall') {
    $i1.Text = '正在卸载'
    $script:lblInstLead.Text = '正在卸载，请稍候…'
  } else {
    $i1.Text = '正在安装'
    $script:lblInstLead.Text = '正在安装，请稍候…'
  }
  $script:lblStage.Text = '准备中…'
  $script:lblElapsed.Text = ''
  $script:progressBar.Value = 0
  $script:btnCancel.Visible = $true
  $script:btnCancel.Text = '取消任务'
  $script:btnCancel.ForeColor = [System.Drawing.ColorTranslator]::FromHtml('#DC2626')
}

function Render-Done {
  $d = $script:lastDone
  if (-not $d) { return }
  $ok = ($d.ExitCode -eq 0) -and (-not $d.Cancelled)
  if ($script:mode -eq 'uninstall') {
    $script:lblDoneTitle.Text = if ($ok) { '卸载完成' } else { '卸载未完成' }
    $script:lblDoneTitle.ForeColor = [System.Drawing.ColorTranslator]::FromHtml($(if ($ok) { '#16A34A' } else { '#DC2626' }))
    $script:lblDoneSub.Text = if ($ok) { '已停止服务、移除自启动、注销 PowerPoint 并清理安装目录。' } elseif ($d.Cancelled) { '卸载已取消。' } else { '卸载过程出现问题，请查看日志。' }
    $script:lblDoneList.Text = ''
    $script:btnPpt.Visible = $false; $script:btnMcp.Visible = $false
    $script:btnDir.Visible = $false; $script:btnShortcut.Visible = $false
    $script:btnViewLog.Visible = (-not $ok)
    $script:btnRetry.Visible = $false
    return
  }
  if ($ok) {
    $env = Get-EnvCheck
    $script:lblDoneTitle.Text = '安装成功'
    $script:lblDoneTitle.ForeColor = [System.Drawing.ColorTranslator]::FromHtml('#16A34A')
    $v = if ($env.InstPresent) { $env.InstVersion } else { '?' }
    $script:lblDoneSub.Text = '版本 v' + $v + '  ·  ' + $d.At + $script:nl + '请完全退出并重新打开 PowerPoint，在功能区「开始」选项卡点击「模板助手」。'
    $lines = @('安装目录：%LOCALAPPDATA%\PPT-AI-Addin')
    if ($env.InstPresent) { $lines += '已安装版本：v' + $env.InstVersion }
    if ($env.RegOK) { $lines += 'PowerPoint 注册：已注册（当前用户）' }
    if ($env.SvcVersion) { $lines += '本地服务：运行中 · v' + $env.SvcVersion + ' · 端口 ' + $env.SvcPort }
    if ($env.Mcp) { $lines += 'MCP：在线' }
    $script:lblDoneList.Text = ($lines -join $script:nl)
    $script:btnPpt.Visible = $true; $script:btnMcp.Visible = $true
    $script:btnDir.Visible = $true; $script:btnShortcut.Visible = $true
    $script:btnViewLog.Visible = $false; $script:btnRetry.Visible = $false
  } else {
    $script:lblDoneTitle.Text = if ($d.Cancelled) { '已取消' } else { '安装失败' }
    $script:lblDoneTitle.ForeColor = [System.Drawing.ColorTranslator]::FromHtml('#DC2626')
    $script:lblDoneSub.Text = if ($d.Cancelled) { '安装任务已被取消。' } else { '安装过程出现问题（退出码 ' + $d.ExitCode + '），请查看日志排查。' }
    $script:lblDoneList.Text = ''
    $script:btnPpt.Visible = $false; $script:btnMcp.Visible = $false
    $script:btnDir.Visible = $false; $script:btnShortcut.Visible = $false
    $script:btnViewLog.Visible = $true; $script:btnRetry.Visible = $true
  }
}

# 定时器：轮询任务进度
$script:timer = New-Object System.Windows.Forms.Timer
$script:timer.Interval = 500
$script:timer.Add_Tick({
  if ($script:step -ne 3) { return }
  if (-not $script:job) { return }
  Sync-JobLog
  Render-NewLogs
  $j = $script:job
  $el = [int](((Get-Date) - $j.Start).TotalSeconds)
  if ($j.TotalSteps -ge 1) { $script:lblElapsed.Text = '已用时 ' + $el + ' 秒 · ' + $j.Step + '/' + $j.TotalSteps } else { $script:lblElapsed.Text = '已用时 ' + $el + ' 秒' }
  if (-not $j.Proc.HasExited) {
    if ($j.StepLabel) { $script:lblStage.Text = $j.StepLabel }
    if ($j.TotalSteps -ge 1) {
      $pct = [math]::Floor(($j.Step / $j.TotalSteps) * 100)
      if ($pct -lt 4) { $pct = 4 }
      if ($pct -le 100) { $script:progressBar.Value = $pct }
    }
  } else {
    $script:timer.Stop()
    Sync-JobLog -FlushPending
    Render-NewLogs
    $exit = $j.Proc.ExitCode
    $script:lastDone = [pscustomobject]@{ Kind = $j.Kind; ExitCode = $exit; Cancelled = $script:cancelled; At = (Get-Date -Format 'HH:mm:ss') }
    $note = if ($script:cancelled) { '已取消' } elseif ($exit -eq 0) { '成功' } else { '失败（退出码 ' + $exit + '）' }
    Add-LogLine ('[任务] ' + $j.Kind + ' → ' + $note + '，耗时 ' + $el + ' 秒')
    if ($exit -eq 0 -and -not $script:cancelled) { $script:progressBar.Value = 100; $script:lblStage.Text = '完成' }
    $script:job = $null
    if ($script:lastTmpBat -and (Test-Path -LiteralPath $script:lastTmpBat)) { Remove-Item -LiteralPath $script:lastTmpBat -Force -ErrorAction SilentlyContinue }
    $script:lastTmpBat = $null
    Set-Step 4
  }
})

# ---- 事件 ----
$script:btnNext.Add_Click({
  if ($script:step -eq 0) { Set-Step 1 }
  elseif ($script:step -eq 1) { if ($script:envOK) { Set-Step 2 } }
  elseif ($script:step -eq 2) {
    if ($script:mode -eq 'install') {
      $args = '--no-pause'
      if ($script:radSkip.Checked) { $args += ' --skip-build' }
      if ($script:chkTests.Checked) { $args += ' --with-tests' }
      if ($script:chkSkipE2e.Checked) { $args += ' --skip-e2e' }
      if (-not (Start-BatJob '一键安装' $batInstall $args)) { return }
    } else {
      if (-not (Start-BatJob '卸载' $batUninstall '/y')) { return }
    }
    Set-Step 3
    $script:timer.Start()
  }
  elseif ($script:step -eq 4) { $form.Close() }
})
$script:btnBack.Add_Click({
  if ($script:step -eq 2) { Set-Step 1 }
  elseif ($script:step -eq 1) { Set-Step 0 }
})
$script:btnCancel.Add_Click({
  if ($script:step -eq 3 -and $script:job -and -not $script:job.Proc.HasExited) {
    if ([System.Windows.Forms.MessageBox]::Show('确定取消当前任务？', '确认', 'YesNo', 'Question') -eq 'Yes') { Cancel-Job }
  } else {
    Quit-Center
  }
})
function Quit-Center {
  if ($script:job -and -not $script:job.Proc.HasExited) {
    [System.Windows.Forms.MessageBox]::Show('任务运行中，请先等待完成或取消。', '提示', 'OK', 'Warning') | Out-Null
    return
  }
  $form.Close()
}
$btnUninstallLink.Add_Click({
  if ([System.Windows.Forms.MessageBox]::Show('进入卸载流程：将停止服务、移除登录自启动、从 PowerPoint 注销加载项，并删除安装目录（%LOCALAPPDATA%\PPT-AI-Addin）。' + $script:nl + '确定继续吗？', '确认', 'YesNo', 'Question') -ne 'Yes') { return }
  $script:mode = 'uninstall'
  $form.Text = '模板助手 · 卸载向导'
  $lblHdrTitle.Text = '模板助手 · 卸载向导'
  $lblHdrSub.Text = '停止服务 · 注销加载项 · 清理文件'
  Set-Step 3
  if (Start-BatJob '卸载' $batUninstall '/y') { $script:timer.Start() }
})
$btnRecheck.Add_Click({ Render-Check })
$btnDlNode.Add_Click({ Open-Target 'nodejs' })
$script:radFull.Add_CheckedChanged({ Sync-Options })
$script:radSkip.Add_CheckedChanged({ Sync-Options })
$script:chkTests.Add_CheckedChanged({ Sync-Options })
$script:chkSkipE2e.Add_CheckedChanged({ Sync-Options })
$script:btnPpt.Add_Click({ Open-Target 'powerpoint' })
$script:btnMcp.Add_Click({
  $e = Start-McpWindow
  if ($e) { [System.Windows.Forms.MessageBox]::Show($e, '提示', 'OK', 'Warning') | Out-Null }
})
$script:btnDir.Add_Click({ Open-Target 'install_dir' })
$script:btnShortcut.Add_Click({
  $e = New-DesktopShortcut
  if ($e) { [System.Windows.Forms.MessageBox]::Show($e, '提示', 'OK', 'Warning') | Out-Null }
  else { [System.Windows.Forms.MessageBox]::Show('已创建桌面快捷方式「模板助手安装向导」。', '完成', 'OK', 'Information') | Out-Null }
})
$script:btnViewLog.Add_Click({ Set-Step 3 })
$script:btnRetry.Add_Click({
  $script:mode = 'install'
  $form.Text = '模板助手 · 安装向导'
  $lblHdrTitle.Text = '模板助手 · 安装向导'
  $lblHdrSub.Text = '环境检测 · 安装 · 自检'
  Set-Step 2
})
$form.Add_FormClosing({
  if ($script:job -and -not $script:job.Proc.HasExited) {
    $_.Cancel = $true
    [System.Windows.Forms.MessageBox]::Show('任务运行中，请先等待完成或取消。', '提示', 'OK', 'Warning') | Out-Null
  }
})

# ---- 启动 ----
Set-Step 0
if ($GuiTest) {
  try {
    $form.Show()
    Write-Output ('ICON: ' + $(if ($form.Icon) { $form.Icon.Width.ToString() + 'x' + $form.Icon.Height.ToString() } else { 'NONE' }))
    # 1) 欢迎 → 环境检查 via real button click
    $script:btnNext.PerformClick()
    if ($script:step -ne 1) { throw 'next click did not advance to step 1 (step=' + $script:step + ')' }
    Write-Output 'CLICK-NEXT-OK'
    # 2) 环境检查 → 安装选项 via button
    $script:btnNext.PerformClick()
    if ($script:step -ne 2) { throw 'next click did not advance to step 2' }
    Write-Output 'CLICK-OPTIONS-OK'
    # 3) 环境检查 back
    $script:btnBack.PerformClick()
    if ($script:step -ne 1) { throw 'back click failed' }
    Write-Output 'CLICK-BACK-OK'
    $script:btnNext.PerformClick()
    # 4) start a mock install job via button handler (step 2 -> 开始安装)
    $tmp = Join-Path $env:TEMP ('ppt-addin-ui-guijob-' + [guid]::NewGuid().ToString('N') + '.bat')
    $content = '@chcp 65001 >nul' + $script:crlf + '@echo off' + $script:crlf + 'echo [0/6] 检查运行环境' + $script:crlf + 'ping -n 3 127.0.0.1 >nul' + $script:crlf + 'echo [4/6] 启动本地服务' + $script:crlf + 'echo [PASS] 环境正常' + $script:crlf + 'exit /b 0'
    [System.IO.File]::WriteAllText($tmp, $content, (New-Object System.Text.UTF8Encoding($false)))
    $script:job = $null
    $null = Start-BatJob 'guijob' $tmp ''
    Set-Step 3
    # simulate timer ticks until job finishes
    for ($i = 0; $i -lt 30; $i++) {
      Start-Sleep -Milliseconds 400
      $tick = $script:timer
      # invoke the tick handler manually
      $script:timer_Tick = $null
      $done = $false
      Sync-JobLog
      Render-NewLogs
      if ($script:job -and $script:job.StepLabel -and $script:job.StepLabel.StartsWith('[')) { throw 'StepLabel keeps raw prefix: ' + $script:job.StepLabel }
      if ($script:job -and $script:job.Proc.HasExited) { $done = $true }
      if ($done) {
        $el = [int](((Get-Date) - $script:job.Start).TotalSeconds)
        $exit = $script:job.Proc.ExitCode
        $script:lastDone = [pscustomobject]@{ Kind = $script:job.Kind; ExitCode = $exit; Cancelled = $false; At = '00:00:00' }
        if ($exit -eq 0) { $script:progressBar.Value = 100 }
        $script:job = $null
        break
      }
    }
    if (-not $script:lastDone) { throw 'mock job did not finish' }
    Write-Output ('JOB-DONE exit=' + $script:lastDone.ExitCode + ' progress=' + $script:progressBar.Value)
    Set-Step 4
    Write-Output ('DONE-TITLE: ' + $script:lblDoneTitle.Text)
    Write-Output ('DONE-BTN-PPT-VISIBLE: ' + $script:btnPpt.Visible)
    $logText = $script:richLog.Text
    if ($logText.IndexOf([char]0xFFFD) -ge 0) { throw 'log box contains mojibake' }
    if ($logText -notmatch '检查运行环境') { throw 'log box missing Chinese lines' }
    if ($logText -notmatch '环境正常') { throw 'log box missing PASS line' }
    Write-Output ('LOG-BOX: ' + $logText.Replace($script:nl, ' | '))
    Write-Output ('LOG-VISIBLE: True')
    Remove-Item -LiteralPath $tmp -Force -ErrorAction SilentlyContinue
  } catch {
    Write-Output ('GUI-ERR: ' + $_.Exception.Message)
    Write-Output ('GUI-ERR-LINE: ' + $_.InvocationInfo.ScriptLineNumber + ' ' + $_.InvocationInfo.Line)
    exit 2
  }
  exit 0
}
$script:timer.Start()
try {
  [System.Windows.Forms.Application]::Run($form)
} catch {
  try { [System.IO.File]::WriteAllText((Join-Path $env:TEMP 'ppt-wizard-crash.log'), ($_.Exception.ToString()), (New-Object System.Text.UTF8Encoding($true))) } catch {}
  throw
}
try { $script:mutex.ReleaseMutex() } catch {}
exit 0
