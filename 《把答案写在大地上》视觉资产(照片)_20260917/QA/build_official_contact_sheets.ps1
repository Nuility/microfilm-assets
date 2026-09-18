param(
    [string]$SourceDir = "D:\lzy\微电影\《把答案写在大地上》视觉资产(照片)_20260917\镜头关键帧",
    [string]$OutputDir = "D:\lzy\微电影\《把答案写在大地上》视觉资产(照片)_20260917\QA\联系表_正式版"
)

Add-Type -AssemblyName System.Drawing
New-Item -ItemType Directory -Force -Path $OutputDir | Out-Null

$cellWidth = 720
$cellHeight = 405
$labelHeight = 42
$cols = 2
$rows = 3

for ($sheet = 0; $sheet -lt 6; $sheet++) {
    $canvas = New-Object System.Drawing.Bitmap ($cellWidth * $cols), (($cellHeight + $labelHeight) * $rows)
    $graphics = [System.Drawing.Graphics]::FromImage($canvas)
    $graphics.Clear([System.Drawing.Color]::FromArgb(24, 24, 24))
    $graphics.InterpolationMode = [System.Drawing.Drawing2D.InterpolationMode]::HighQualityBicubic
    $font = New-Object System.Drawing.Font('Arial', 22, [System.Drawing.FontStyle]::Bold)
    $brush = [System.Drawing.Brushes]::White

    for ($slot = 0; $slot -lt 6; $slot++) {
        $index = $sheet * 6 + $slot + 1
        $v02Name = 'KF_{0:D2}_v02.png' -f $index
        $v01Name = 'KF_{0:D2}_v01.png' -f $index
        $v02Path = Join-Path $SourceDir $v02Name
        $v01Path = Join-Path $SourceDir $v01Name

        if (Test-Path -LiteralPath $v02Path) {
            $path = $v02Path
            $version = 'v02'
        }
        else {
            $path = $v01Path
            $version = 'v01'
        }

        $col = $slot % $cols
        $row = [Math]::Floor($slot / $cols)
        $x = $col * $cellWidth
        $y = $row * ($cellHeight + $labelHeight)
        $graphics.DrawString(('KF{0:D2} {1}' -f $index, $version), $font, $brush, $x + 12, $y + 7)

        if (Test-Path -LiteralPath $path) {
            $image = [System.Drawing.Image]::FromFile($path)
            $graphics.DrawImage($image, $x, $y + $labelHeight, $cellWidth, $cellHeight)
            $image.Dispose()
        }
    }

    $outPath = Join-Path $OutputDir ('OFFICIAL_{0:D2}-{1:D2}.jpg' -f ($sheet * 6 + 1), ($sheet * 6 + 6))
    $canvas.Save($outPath, [System.Drawing.Imaging.ImageFormat]::Jpeg)
    $font.Dispose()
    $graphics.Dispose()
    $canvas.Dispose()
}
