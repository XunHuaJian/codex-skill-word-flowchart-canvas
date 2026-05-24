param(
  [Parameter(Mandatory = $true)]
  [string]$OutPath
)

$ErrorActionPreference = 'Stop'

$outDir = Split-Path -Parent $OutPath
if ($outDir -and -not (Test-Path -LiteralPath $outDir)) {
  New-Item -ItemType Directory -Path $outDir | Out-Null
}

$word = New-Object -ComObject Word.Application
$word.Visible = $false
$word.DisplayAlerts = 0

try {
  $doc = $word.Documents.Add()
  $doc.PageSetup.TopMargin = 18
  $doc.PageSetup.BottomMargin = 18
  $doc.PageSetup.LeftMargin = 18
  $doc.PageSetup.RightMargin = 18

  $items = $doc.Shapes.AddCanvas(18, 18, 555, 785, $doc.Range(0, 0)).CanvasItems

  $shapeSpecs = @(
    @(5, 165, 8, 130, 30, '开始', 12),
    @(5, 130, 52, 200, 30, '系统初始化', 12),
    @(5, 100, 96, 260, 30, '采集温湿度数据', 12),
    @(5, 80, 140, 300, 30, '检测红外/柜门状态', 12),
    @(4, 75, 190, 310, 70, '判断湿度是否超限', 12),
    @(5, 445, 209, 105, 32, '启动换气除湿', 10),
    @(5, 90, 280, 280, 32, '保持或关闭换气除湿', 11),
    @(4, 45, 340, 370, 82, '判断是否有人靠近或开门', 12),
    @(5, 445, 356, 105, 52, "开启照明`r禁止消毒", 10),
    @(5, 85, 450, 290, 32, '关闭或延时关闭照明', 11),
    @(4, 40, 520, 380, 84, '判断是否收到本地/远程指令', 11),
    @(5, 445, 546, 105, 32, '解析并执行指令', 9),
    @(5, 140, 625, 180, 32, '继续自动控制', 11),
    @(5, 110, 690, 240, 32, '更新显示并上传数据', 11),
    @(5, 120, 740, 220, 32, '返回循环检测', 11)
  )

  foreach ($sp in $shapeSpecs) {
    $s = $items.AddShape($sp[0], $sp[1], $sp[2], $sp[3], $sp[4])
    $s.TextFrame.TextRange.Text = $sp[5]
    $s.TextFrame.TextRange.Font.NameFarEast = 'Microsoft YaHei'
    $s.TextFrame.TextRange.Font.Size = $sp[6]
    $s.TextFrame.TextRange.ParagraphFormat.Alignment = 1
    $s.TextFrame.VerticalAnchor = 3
    $s.Fill.ForeColor.RGB = 16777215
    $s.Line.ForeColor.RGB = 3355443
    $s.Line.Weight = 1.25
  }

  $labels = @(
    @(405, 206, '是'),
    @(245, 262, '否'),
    @(422, 358, '是'),
    @(245, 426, '否'),
    @(422, 540, '是'),
    @(245, 607, '否')
  )

  foreach ($t in $labels) {
    $tb = $items.AddTextbox(1, $t[0], $t[1], 24, 18)
    $tb.TextFrame.TextRange.Text = $t[2]
    $tb.Fill.Visible = 0
    $tb.Line.Visible = 0
    $tb.TextFrame.TextRange.Font.NameFarEast = 'Microsoft YaHei'
    $tb.TextFrame.TextRange.Font.Size = 10
    $tb.TextFrame.TextRange.ParagraphFormat.Alignment = 1
  }

  function Add-Connector([double]$x1, [double]$y1, [double]$x2, [double]$y2, [bool]$arrow) {
    $c = $items.AddConnector(1, $x1, $y1, $x2, $y2)
    $c.Line.ForeColor.RGB = 3355443
    $c.Line.Weight = 1.25
    if ($arrow) {
      $c.Line.EndArrowheadStyle = 3
    } else {
      $c.Line.EndArrowheadStyle = 1
    }
  }

  Add-Connector 230 38 230 52 $true
  Add-Connector 230 82 230 96 $true
  Add-Connector 230 126 230 140 $true
  Add-Connector 230 170 230 190 $true

  Add-Connector 385 225 445 225 $true
  Add-Connector 230 260 230 280 $true
  Add-Connector 230 312 230 340 $true
  Add-Connector 497 241 497 325 $false
  Add-Connector 497 325 230 325 $false
  Add-Connector 230 325 230 340 $true

  Add-Connector 415 381 445 381 $true
  Add-Connector 230 422 230 450 $true
  Add-Connector 230 482 230 520 $true
  Add-Connector 497 408 497 502 $false
  Add-Connector 497 502 230 502 $false
  Add-Connector 230 502 230 520 $true

  Add-Connector 420 562 445 562 $true
  Add-Connector 230 604 230 625 $true
  Add-Connector 230 657 230 690 $true
  Add-Connector 497 578 497 706 $false
  Add-Connector 497 706 350 706 $true

  Add-Connector 230 722 230 740 $true

  $doc.SaveAs2($OutPath, 12)
  $doc.Saved = $true
  $doc.Close(0)
  Write-Output $OutPath
} finally {
  $word.Quit()
}
