param(
  [string]$ProcessName = 'ppt-ai-addin',
  [string]$RuntimeFile = (Join-Path $env:APPDATA 'ppt-ai-addin\runtime.json'),
  [int]$TimeoutSeconds = 15
)

$ErrorActionPreference = 'SilentlyContinue'

$port = $null
if (Test-Path -LiteralPath $RuntimeFile) {
  try {
    $runtime = Get-Content -Raw -LiteralPath $RuntimeFile | ConvertFrom-Json
    if ($runtime.port) { $port = [int]$runtime.port }
  } catch {
    $port = $null
  }
}

function Get-TargetProcess {
  $targets = @()
  $targets += Get-Process -Name $ProcessName
  if ($port) {
    $targets += Get-NetTCPConnection -LocalPort $port -State Listen | ForEach-Object {
      Get-Process -Id $_.OwningProcess
    }
  }
  $targets | Where-Object { $_ } | Sort-Object Id -Unique
}

$deadline = (Get-Date).AddSeconds($TimeoutSeconds)
do {
  $targets = @(Get-TargetProcess)
  if ($targets.Count -eq 0) { exit 0 }

  foreach ($p in $targets) {
    Stop-Process -Id $p.Id -Force
  }

  Start-Sleep -Milliseconds 500
} while ((Get-Date) -lt $deadline)

$left = @(Get-TargetProcess)
if ($left.Count -gt 0) {
  $ids = ($left | ForEach-Object { $_.Id }) -join ', '
  Write-Host "[ERROR] Stale service is still running. PID: $ids. Close PowerPoint or restart Windows, then retry."
  exit 1
}

exit 0
