param(
    [string]$RootDir = "D:\lzy\微电影\《把答案写在大地上》视觉资产(照片)_20260917",
    [string]$OutputDir = "D:\lzy\微电影\《把答案写在大地上》视觉资产(照片)_20260917\QA\视觉锚点总览"
)

Add-Type -AssemblyName System.Drawing
New-Item -ItemType Directory -Force -Path $OutputDir | Out-Null

$cellWidth = 720
$cellHeight = 405
$labelHeight = 42
$cols = 2
$rows = 3

function New-ReferenceBoard {
    param(
        [string]$FileName,
        [array]$Items
    )

    $canvas = New-Object System.Drawing.Bitmap ($cellWidth * $cols), (($cellHeight + $labelHeight) * $rows)
    $graphics = [System.Drawing.Graphics]::FromImage($canvas)
    $graphics.Clear([System.Drawing.Color]::FromArgb(24, 24, 24))
    $graphics.InterpolationMode = [System.Drawing.Drawing2D.InterpolationMode]::HighQualityBicubic
    $font = New-Object System.Drawing.Font('Arial', 22, [System.Drawing.FontStyle]::Bold)
    $brush = [System.Drawing.Brushes]::White

    for ($slot = 0; $slot -lt $Items.Count; $slot++) {
        $item = $Items[$slot]
        $col = $slot % $cols
        $row = [Math]::Floor($slot / $cols)
        $x = $col * $cellWidth
        $y = $row * ($cellHeight + $labelHeight)
        $graphics.DrawString($item.Label, $font, $brush, $x + 12, $y + 7)
        $path = Join-Path $RootDir $item.Path
        if (Test-Path -LiteralPath $path) {
            $source = [System.Drawing.Image]::FromFile($path)
            $graphics.DrawImage($source, $x, $y + $labelHeight, $cellWidth, $cellHeight)
            $source.Dispose()
        }
    }

    $outputPath = Join-Path $OutputDir $FileName
    $canvas.Save($outputPath, [System.Drawing.Imaging.ImageFormat]::Jpeg)
    $font.Dispose()
    $graphics.Dispose()
    $canvas.Dispose()
}

$identityItems = @(
    @{ Label = 'A01 LIN YUE - CHAR v02'; Path = '角色设定\CHAR_A01_林悦_v02.png' },
    @{ Label = 'A02 TEACHER CHEN - CHAR v02'; Path = '角色设定\CHAR_A02_陈老师_v02.png' },
    @{ Label = 'A03 ZHOU JIANGUO - CHAR v02'; Path = '角色设定\CHAR_A03_周建国_v02.png' },
    @{ Label = 'P03 LANDSCAPE PHONE GIMBAL'; Path = '道具设定\PROP_P03_手机稳定器_v01.png' },
    @{ Label = 'P06 OLD BRASS KEY'; Path = '道具设定\PROP_P06_旧钥匙_v01.png' }
)

$sceneItems = @(
    @{ Label = 'S01/S05 CLASSROOM - KF01'; Path = '镜头关键帧\KF_01_v01.png' },
    @{ Label = 'S02 CORRIDOR - KF08'; Path = '镜头关键帧\KF_08_v01.png' },
    @{ Label = 'S03 ALLEY - KF10'; Path = '镜头关键帧\KF_10_v02.png' },
    @{ Label = 'S03 HISTORY ROOM - KF13'; Path = '镜头关键帧\KF_13_v02.png' },
    @{ Label = 'S04 COURTYARD - KF23'; Path = '镜头关键帧\KF_23_v02.png' }
)

New-ReferenceBoard -FileName 'REFERENCE_角色与道具.jpg' -Items $identityItems
New-ReferenceBoard -FileName 'REFERENCE_场景空间.jpg' -Items $sceneItems
