# 把生成的单页 pptx 插入当前 PowerPoint 演示文稿（当前选中页之后）
# 用法：powershell -NoProfile -ExecutionPolicy Bypass -File insert.ps1 -PptxPath <文件>
param([string]$PptxPath)
$ErrorActionPreference = 'Stop'
# 输出统一 UTF-8（否则 Node 端按 UTF-8 解码 GBK 会乱码）
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8
if (-not (Test-Path $PptxPath)) { Write-Error "文件不存在: $PptxPath"; exit 1 }

$ppt = $null
try {
  # 优先连接正在运行的 PowerPoint 实例（GetActiveObject）
  $ppt = [System.Runtime.InteropServices.Marshal]::GetActiveObject('PowerPoint.Application')
} catch {
  try { $ppt = New-Object -ComObject PowerPoint.Application } catch {
    Write-Error '无法连接 PowerPoint（COM 不可用）。请确认 PowerPoint 已安装且未以管理员身份运行（权限不一致时 COM 拒绝访问）。'
    exit 2
  }
}
try { $ppt.Visible = $true } catch { }

$pres = $null
try { $pres = $ppt.ActivePresentation } catch { }
if (-not $pres) {
  Write-Error '没有打开的演示文稿，请先在 PowerPoint 中打开目标文档后再试。'
  exit 3
}

# 插入位置：当前选中幻灯片之后；无选中时插到末尾。
# 注意 Slides.InsertFromFile 的 Index 语义 = 「在第 Index 页之后插入」，合法范围 0..Count
# （实测 Count=40 时传 41 会报 Integer out of range: 41 is not in valid range 0..40），
# 因此末尾追加必须传 Count 而不是 Count+1；选中页 n 直接传 n 即插到其后。
$index = $pres.Slides.Count
try {
  if ($ppt.ActiveWindow.Selection.Type -eq 3) { # ppSelectionSlides
    $index = $ppt.ActiveWindow.Selection.SlideRange.Item(1).SlideIndex
  }
} catch { }

$pres.Slides.InsertFromFile($PptxPath, $index) | Out-Null
Write-Output "已插入到第 $index 页之后（共 $($pres.Slides.Count) 页）"
exit 0
