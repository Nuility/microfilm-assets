param(
    [Parameter(Mandatory = $true)]
    [string]$ProjectRoot,

    [int]$Columns = 2,
    [int]$Rows = 3
)

$resolvedRoot = (Resolve-Path -LiteralPath $ProjectRoot).Path
$indexPath = Join-Path $resolvedRoot '09_正式资产索引.json'
if (-not (Test-Path -LiteralPath $indexPath -PathType Leaf)) {
    throw "找不到正式资产索引：$indexPath"
}

$index = Get-Content -Raw -LiteralPath $indexPath | ConvertFrom-Json
$frames = @($index.official_keyframes) | Where-Object {
    -not [string]::IsNullOrWhiteSpace([string]$_.file) -and
    (Test-Path -LiteralPath (Join-Path $resolvedRoot ([string]$_.file)) -PathType Leaf)
}

if ($frames.Count -eq 0) {
    throw '正式资产索引中没有可读取的关键帧。'
}

$relativeOutputDir = [string]$index.qa.contact_sheet_dir
if ([string]::IsNullOrWhiteSpace($relativeOutputDir)) {
    $relativeOutputDir = 'qa/contact_sheets'
}
$outputDir = Join-Path $resolvedRoot $relativeOutputDir
New-Item -ItemType Directory -Force -Path $outputDir | Out-Null

Add-Type -AssemblyName System.Drawing
$cellWidth = 720
$cellHeight = 405
$labelHeight = 42
$perSheet = $Columns * $Rows
$sheetCount = [Math]::Ceiling($frames.Count / [double]$perSheet)

for ($sheet = 0; $sheet -lt $sheetCount; $sheet++) {
    $canvas = New-Object System.Drawing.Bitmap ($cellWidth * $Columns), (($cellHeight + $labelHeight) * $Rows)
    $graphics = [System.Drawing.Graphics]::FromImage($canvas)
    $graphics.Clear([System.Drawing.Color]::FromArgb(24, 24, 24))
    $graphics.InterpolationMode = [System.Drawing.Drawing2D.InterpolationMode]::HighQualityBicubic
    $font = New-Object System.Drawing.Font('Arial', 22, [System.Drawing.FontStyle]::Bold)
    $brush = [System.Drawing.Brushes]::White

    for ($slot = 0; $slot -lt $perSheet; $slot++) {
        $frameIndex = $sheet * $perSheet + $slot
        if ($frameIndex -ge $frames.Count) { break }
        $frame = $frames[$frameIndex]
        $col = $slot % $Columns
        $row = [Math]::Floor($slot / $Columns)
        $x = $col * $cellWidth
        $y = $row * ($cellHeight + $labelHeight)
        $label = "$($frame.id) $($frame.version)"
        $graphics.DrawString($label, $font, $brush, $x + 12, $y + 7)

        $path = Join-Path $resolvedRoot ([string]$frame.file)
        $source = [System.Drawing.Image]::FromFile($path)
        $graphics.DrawImage($source, $x, $y + $labelHeight, $cellWidth, $cellHeight)
        $source.Dispose()
    }

    $startNumber = $sheet * $perSheet + 1
    $endNumber = [Math]::Min(($sheet + 1) * $perSheet, $frames.Count)
    $outPath = Join-Path $outputDir ('OFFICIAL_{0:D3}-{1:D3}.jpg' -f $startNumber, $endNumber)
    $canvas.Save($outPath, [System.Drawing.Imaging.ImageFormat]::Jpeg)
    $font.Dispose()
    $graphics.Dispose()
    $canvas.Dispose()
}

Write-Host "已生成 $sheetCount 张联系表：$outputDir"
