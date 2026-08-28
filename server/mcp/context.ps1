# 读取当前 PowerPoint 演示文稿的结构化上下文（不读取/传输 PPTX 二进制）
# 用法：powershell -NoProfile -ExecutionPolicy Bypass -File context.ps1 -Action <presentation|current-slide|slide|inspect> [-Index n] [-Id n] -OutFile <json路径>
# 结果写入 OutFile（UTF-8 JSON），避免 stdout 编码/协议污染。
param(
  [string]$Action = '',
  [string]$Index = '',
  [string]$Id = '',
  [string]$OutFile = ''
)
$ErrorActionPreference = 'Stop'
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8

function Emit($obj) {
  $json = $obj | ConvertTo-Json -Depth 20
  $json | Out-File -FilePath $OutFile -Encoding utf8
}

# 形状类型数字 → 名称（MsoShapeType；PowerPoint 专有类型已含）
$SHAPE_TYPES = @{
  1='autoShape'; 2='callout'; 3='chart'; 4='comment'; 5='freeform'; 6='group';
  7='embeddedOle'; 8='formControl'; 9='line'; 10='linkedOle'; 11='linkedPicture';
  12='ole'; 13='picture'; 14='placeholder'; 15='textEffect'; 16='media'; 17='textBox';
  18='scriptAnchor'; 19='table'; 20='canvas'; 21='diagram'; 22='ink'; 23='inkComment';
  24='smartArt'; 25='slicer'; 26='webVideo'; 27='graphic'; 28='contentApp'; 29='model3D'
}

$MAX_TEXT = 300      # 单元素文本最长字符数（裁剪）
$MAX_SHAPES = 300    # 单页形状读取上限
$MAX_GROUP_DEPTH = 2 # 组合对象递归深度上限
$MAX_SLIDES = 500    # 页面列表上限

# 读取单个形状（单位换算：磅 → 英寸；文本/字体/图片/表格/组合标记）
function Read-Shape($s, $depth) {
  $o = [ordered]@{}
  $typeNum = 0
  try { $typeNum = [int]$s.Type } catch { }
  $o.id = $s.Id
  $o.name = $s.Name
  $o.type = $SHAPE_TYPES[$typeNum]
  if (-not $o.type) { $o.type = 'unknown(' + $typeNum + ')' }
  $o.left = [math]::Round([double]$s.Left / 72.0, 3)
  $o.top = [math]::Round([double]$s.Top / 72.0, 3)
  $o.width = [math]::Round([double]$s.Width / 72.0, 3)
  $o.height = [math]::Round([double]$s.Height / 72.0, 3)
  try { $o.rotation = [math]::Round([double]$s.Rotation, 1) } catch { $o.rotation = 0 }
  $o.isZeroSize = ([double]$s.Width -eq 0 -or [double]$s.Height -eq 0)
  $o.isTextBox = $false
  $o.isPicture = ($typeNum -eq 13 -or $typeNum -eq 11)
  $o.isTable = $false
  $o.isGroup = ($typeNum -eq 6)
  try { $o.isTextBox = [bool]$s.HasTextFrame } catch { }
  try { $o.isTable = [bool]$s.HasTable } catch { }
  if ($s.HasTable) { try { $o.tableInfo = [ordered]@{ rows = $s.Table.Rows.Count; cols = $s.Table.Columns.Count } } catch { } }

  # 文本（含字体信息，只取首段/首字符字体作为代表）
  $text = ''
  try {
    if ($s.HasTextFrame -and $s.TextFrame.HasText) {
      $text = [string]$s.TextFrame.TextRange.Text
      $text = $text -replace [char]13, "
"
      if ($text.Length -gt $MAX_TEXT) { $text = $text.Substring(0, $MAX_TEXT) + '…(已截断)' }
      $o.text = $text
      try {
        $f = $s.TextFrame.TextRange.Font
        $color = ''
        try { $color = ('{0:X6}' -f ([int]($f.Color.RGB) -band 0xFFFFFF)) } catch { }
        $o.font = [ordered]@{ name = [string]$f.Name; size = [double]$f.Size; bold = [bool]$f.Bold; italic = [bool]$f.Italic; color = $color }
      } catch { }
    }
  } catch { }

  # 组合对象：递归子形状（有限深度）
  if ($typeNum -eq 6 -and $depth -lt $MAX_GROUP_DEPTH) {
    $kids = @()
    $kidCount = 0
    foreach ($c in $s.GroupItems) {
      $kidCount++
      if ($kidCount -gt $MAX_SHAPES) { break }
      $kids += Read-Shape $c ($depth + 1)
    }
    $o.childCount = $kidCount
    $o.children = $kids
  }
  return $o
}

# 读取一页（$full：完整形状列表；否则只给计数）
function Read-Slide($slide, $full) {
  $o = [ordered]@{}
  $o.index = [int]$slide.SlideIndex
  $o.slideId = [int64]$slide.SlideID
  try { $o.layoutName = [string]$slide.CustomLayout.Name } catch { }
  $total = 0
  try { $total = [int]$slide.Shapes.Count } catch { }
  $o.shapeCount = $total
  if ($full) {
    $shapes = @()
    $read = 0
    $truncated = $false
    foreach ($sh in $slide.Shapes) {
      $read++
      if ($read -gt $MAX_SHAPES) { $truncated = $true; break }
      $shapes += Read-Shape $sh 0
    }
    $o.truncated = $truncated
    $o.shapes = $shapes
  }
  return $o
}

# 当前查看/选中页（无选中页时回退到视图当前页）
function Get-CurrentSlide {
  $sel = $ppt.ActiveWindow.Selection
  $selType = 0
  try { $selType = [int]$sel.Type } catch { }
  if ($selType -eq 2) { # ppSelectionSlides
    return $sel.SlideRange.Item(1)
  }
  # ppSelectionNone / ppSelectionShapes / 其他 → 视图当前页
  return $ppt.ActiveWindow.View.Slide
}

# ---------- 主流程 ----------
$ppt = $null
try {
  $ppt = [System.Runtime.InteropServices.Marshal]::GetActiveObject('PowerPoint.Application')
} catch {
  Emit @{ ok = $false; error = 'PowerPoint 未在运行。请先打开 PowerPoint 并打开目标演示文稿，再调用本工具。' }
  exit 0
}
$pres = $null
try { $pres = $ppt.ActivePresentation } catch { }
if (-not $pres) {
  Emit @{ ok = $false; error = 'PowerPoint 正在运行但没有打开的演示文稿。请先打开目标文档。' }
  exit 0
}

switch ($Action) {
  'doc-path' {
    # 返回当前文档的磁盘路径（供后端直读文件，绕开 Office.js 慢通道）
    $fullName = ''
    try { $fullName = [string]$pres.FullName } catch { }
    if (-not $fullName) {
      Emit @{ ok = $false; error = '无法获取当前文档路径（文档可能尚未保存到磁盘）' }
    } else {
      Emit @{ ok = $true; path = $fullName }
    }
  }
  'presentation' {
    $p = [ordered]@{ ok = $true; kind = 'presentation'; capturedAt = (Get-Date).ToString('o') }
    $p.name = [string]$pres.Name
    $saved = $false
    try { $saved = [bool]$pres.Saved } catch { }
    $p.saved = $saved
    $fullName = ''
    $docPath = ''
    try { $fullName = [string]$pres.FullName } catch { }
    try { $docPath = [string]$pres.Path } catch { }
    $p.fullName = $fullName
    $p.path = $docPath
    $p.slideCount = [int]$pres.Slides.Count
    $p.slideWidthIn = [math]::Round([double]$pres.PageSetup.SlideWidth / 72.0, 3)
    $p.slideHeightIn = [math]::Round([double]$pres.PageSetup.SlideHeight / 72.0, 3)
    $p.slideSize = [ordered]@{ width = $p.slideWidthIn; height = $p.slideHeightIn }

    # 当前选中页 + 当前查看页
    $selType = 0
    try { $selType = [int]$ppt.ActiveWindow.Selection.Type } catch { }
    $cur = Get-CurrentSlide
    if ($cur) {
      $p.currentSlide = [ordered]@{ index = [int]$cur.SlideIndex; slideId = [int64]$cur.SlideID; selectionType = $selType }
    }
    # 当前选中元素（用户选中了形状时）
    $selShapes = @()
    if ($selType -eq 3) { # ppSelectionShapes
      foreach ($sh in $ppt.ActiveWindow.Selection.ShapeRange) {
        $o = [ordered]@{ id = $sh.Id; name = $sh.Name; type = $SHAPE_TYPES[[int]$sh.Type] }
        $txt = ''
        try { if ($sh.HasTextFrame -and $sh.TextFrame.HasText) { $txt = ([string]$sh.TextFrame.TextRange.Text) } } catch { }
        if ($txt.Length -gt 80) { $txt = $txt.Substring(0, 80) + '…' }
        $o.text = $txt
        $selShapes += $o
      }
      $p.selectedShapes = $selShapes
    }
    # 页面列表（index + 稳定 slideId + 元素数；裁剪上限）
    $slides = @()
    $idx = 0
    $slideTrunc = $false
    foreach ($sl in $pres.Slides) {
      $idx++
      if ($idx -gt $MAX_SLIDES) { $slideTrunc = $true; break }
      $slides += Read-Slide $sl $false
    }
    $p.slideListTruncated = $slideTrunc
    $p.slides = $slides
    Emit $p
  }
  'current-slide' {
    $sl = Get-CurrentSlide
    if (-not $sl) {
      Emit @{ ok = $false; error = '无法确定当前页（没有打开任何幻灯片）。' }
    } else {
      $o = Read-Slide $sl $true
      $o.ok = $true
      $o.kind = 'current-slide'
      $o.capturedAt = (Get-Date).ToString('o')
      Emit $o
    }
  }
  'slide' {
    $sl = $null
    if ($Id -ne '') { $sl = $pres.Slides.FindBySlideID([int]$Id) }  # SlideID（文档内稳定 ID）
    elseif ($Index -ne '') { $sl = $pres.Slides.Item([int]$Index) }  # 页码（1 起）
    if (-not $sl) {
      Emit @{ ok = $false; error = '未指定页码/slideId（-Index 或 -Id），或找不到该 slideId。' }
    } else {
      $o = Read-Slide $sl $true
      $o.ok = $true
      $o.kind = 'slide'
      $o.capturedAt = (Get-Date).ToString('o')
      Emit $o
    }
  }
  'inspect' {
    $sl = $null
    if ($Id -ne '') { $sl = $pres.Slides.FindBySlideID([int]$Id) }
    elseif ($Index -ne '') { $sl = $pres.Slides.Item([int]$Index) }
    else { $sl = Get-CurrentSlide }
    if (-not $sl) {
      Emit @{ ok = $false; error = '无法确定要检查的页面（未指定 index/slideId 且没有当前页）。' }
    } else {
      $o = [ordered]@{ ok = $true; kind = 'inspect'; capturedAt = (Get-Date).ToString('o') }
      $o.index = [int]$sl.SlideIndex
      $o.slideId = [int64]$sl.SlideID
      $lines = @()
      $n = 0
      foreach ($sh in $sl.Shapes) {
        $n++
        if ($n -gt $MAX_SHAPES) { $o.truncated = $true; break }
        $line = [ordered]@{}
        $line.id = $sh.Id
        $line.name = [string]$sh.Name
        $line.type = $SHAPE_TYPES[[int]$sh.Type]
        $txt = ''
        try { if ($sh.HasTextFrame -and $sh.TextFrame.HasText) { $txt = ([string]$sh.TextFrame.TextRange.Text) } } catch { }
        $txt = $txt -replace [char]13, " "
        if ($txt.Length -gt 60) { $txt = $txt.Substring(0, 60) + '…' }
        $line.text = $txt
        $line.left = [math]::Round([double]$sh.Left / 72.0, 2)
        $line.top = [math]::Round([double]$sh.Top / 72.0, 2)
        $line.width = [math]::Round([double]$sh.Width / 72.0, 2)
        $line.height = [math]::Round([double]$sh.Height / 72.0, 2)
        $lines += $line
      }
      $o.shapeCount = $n
      $o.shapes = $lines
      Emit $o
    }
  }
  default {
    Emit @{ ok = $false; error = '未知 Action: ' + $Action }
  }
}
exit 0
