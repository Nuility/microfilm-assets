param(
    [Parameter(Mandatory = $true)]
    [string]$ProjectRoot
)

$resolvedRoot = (Resolve-Path -LiteralPath $ProjectRoot).Path
$indexPath = Join-Path $resolvedRoot '09_正式资产索引.json'
if (-not (Test-Path -LiteralPath $indexPath -PathType Leaf)) {
    throw "找不到正式资产索引：$indexPath"
}

try {
    $index = Get-Content -Raw -LiteralPath $indexPath | ConvertFrom-Json
}
catch {
    throw "正式资产索引不是有效JSON：$($_.Exception.Message)"
}

Add-Type -AssemblyName System.Drawing
$missing = New-Object System.Collections.Generic.List[string]
$invalid = New-Object System.Collections.Generic.List[string]
$checked = 0

$entries = @()
$entries += @($index.characters)
$entries += @($index.props)
$entries += @($index.scene_masters)
$entries += @($index.official_keyframes)

$expectedWidth = [int]$index.spec.width
$expectedHeight = [int]$index.spec.height
$targetRatio = 16.0 / 9.0

foreach ($entry in $entries) {
    $relativePath = [string]$entry.file
    $entryId = [string]$entry.id
    if ([string]::IsNullOrWhiteSpace($relativePath)) {
        $missing.Add("$entryId：文件路径为空")
        continue
    }

    $fullPath = Join-Path $resolvedRoot $relativePath
    if (-not (Test-Path -LiteralPath $fullPath -PathType Leaf)) {
        $missing.Add("$entryId：$relativePath")
        continue
    }

    $checked++
    $extension = [System.IO.Path]::GetExtension($fullPath).ToLowerInvariant()
    if ($extension -notin @('.png', '.jpg', '.jpeg')) {
        continue
    }

    try {
        $image = [System.Drawing.Image]::FromFile($fullPath)
        $width = $image.Width
        $height = $image.Height
        $image.Dispose()

        if ($expectedWidth -gt 0 -and $expectedHeight -gt 0) {
            if ($width -ne $expectedWidth -or $height -ne $expectedHeight) {
                $invalid.Add("$entryId：尺寸为${width}x${height}，预期${expectedWidth}x${expectedHeight}")
            }
        }
        elseif ([Math]::Abs(($width / [double]$height) - $targetRatio) -gt 0.02) {
            $invalid.Add("$entryId：${width}x${height}不是16:9")
        }
    }
    catch {
        $invalid.Add("$entryId：无法读取图片 $relativePath")
    }
}

Write-Host "项目：$($index.project)"
Write-Host "已检查文件：$checked"
Write-Host "缺失或未填写：$($missing.Count)"
Write-Host "格式或尺寸异常：$($invalid.Count)"

if ($missing.Count -gt 0) {
    Write-Host "`n缺失清单："
    $missing | ForEach-Object { Write-Host "- $_" }
}

if ($invalid.Count -gt 0) {
    Write-Host "`n异常清单："
    $invalid | ForEach-Object { Write-Host "- $_" }
}

if ($missing.Count -gt 0 -or $invalid.Count -gt 0) {
    exit 1
}

Write-Host "校验通过。"
