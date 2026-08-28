# 黄金路径 E2E 的 PowerPoint COM 打开检查：
# 创建隐藏 PowerPoint 实例 → 打开生成的 pptx（只读）→ 逐页输出 { 形状数/图片数/表格数/文本列表 } → 关闭。
# 用法：powershell -NoProfile -ExecutionPolicy Bypass -File open-check.ps1 -PptxPath <文件> -OutFile <json>
# 退出码：0 = 打开成功（结果写入 OutFile）；2 = PowerPoint COM 不可用（E2E 降级为文件级检查）；
#         3 = 打开失败（文件损坏等）。
param([string]$PptxPath, [string]$OutFile)
$ErrorActionPreference = 'Stop'
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8
if (-not (Test-Path $PptxPath)) { Write-Error "文件不存在: $PptxPath"; exit 1 }

$ppt = $null
try { $ppt = New-Object -ComObject PowerPoint.Application } catch {
  Write-Error 'PowerPoint COM 不可用（未安装 PowerPoint，或以管理员身份运行导致权限冲突）'
  exit 2
}
try { $ppt.Visible = $false } catch { }

$pres = $null
try { $pres = $ppt.Presentations.Open($PptxPath, $true, $false, $false) } catch {
  # 部分 Office 版本对 WithWindow=$false 有限制 → 降级为带窗口打开（实例本身隐藏）
  try { $pres = $ppt.Presentations.Open($PptxPath, $true, $false, $true) } catch {
    Write-Error ('打开失败: ' + $_.Exception.Message)
    try { $ppt.Quit() } catch { }
    exit 3
  }
}

$slidesOut = @()
foreach ($slide in $pres.Slides) {
  $texts = @()
  $pictures = 0
  $tables = 0
  foreach ($sh in $slide.Shapes) {
    try { if ($sh.HasTable) { $tables++ } } catch { }
    $t = 0
    try { $t = [int]$sh.Type } catch { }
    if ($t -eq 13 -or $t -eq 11) { $pictures++ }   # msoPicture / msoLinkedPicture
    try {
      if ($sh.HasTextFrame -and $sh.TextFrame.HasText) {
        $texts += [string]$sh.TextFrame.TextRange.Text
      }
    } catch { }
  }
  $slidesOut += [ordered]@{
    index      = [int]$slide.SlideIndex
    shapeCount = [int]$slide.Shapes.Count
    pictures   = $pictures
    tables     = $tables
    texts      = $texts
  }
}

$out = [ordered]@{ ok = $true; slides = $slidesOut.Count; shapes = $slidesOut } | ConvertTo-Json -Depth 12
$out | Out-File -FilePath $OutFile -Encoding utf8

try { $pres.Close() } catch { }
try { $ppt.Quit() } catch { }
exit 0
