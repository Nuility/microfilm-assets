param(
    [Parameter(Mandatory = $true)]
    [string]$ProjectName,

    [Parameter(Mandatory = $true)]
    [string]$DestinationRoot
)

$templateRoot = Join-Path (Split-Path -Parent $PSScriptRoot) 'project_template'
if (-not (Test-Path -LiteralPath $templateRoot -PathType Container)) {
    throw "找不到项目模板目录：$templateRoot"
}

if (-not (Test-Path -LiteralPath $DestinationRoot -PathType Container)) {
    New-Item -ItemType Directory -Path $DestinationRoot -Force | Out-Null
}

$resolvedRoot = (Resolve-Path -LiteralPath $DestinationRoot).Path
$safeName = $ProjectName.Trim()
if ([string]::IsNullOrWhiteSpace($safeName)) {
    throw '项目名称不能为空。'
}

foreach ($invalidChar in [System.IO.Path]::GetInvalidFileNameChars()) {
    $safeName = $safeName.Replace([string]$invalidChar, '_')
}

$target = Join-Path $resolvedRoot ($safeName + '_视觉资产')
if (Test-Path -LiteralPath $target) {
    throw "目标目录已存在，不会覆盖：$target"
}

Copy-Item -LiteralPath $templateRoot -Destination $target -Recurse

$utf8NoBom = New-Object System.Text.UTF8Encoding($false)
Get-ChildItem -LiteralPath $target -Recurse -File | Where-Object {
    $_.Extension -in @('.md', '.json', '.txt')
} | ForEach-Object {
    $content = [System.IO.File]::ReadAllText($_.FullName)
    $content = $content.Replace('{{PROJECT_NAME}}', $ProjectName)
    [System.IO.File]::WriteAllText($_.FullName, $content, $utf8NoBom)
}

Write-Host "已创建项目：$target"
Write-Host "请从 00_项目简报.md 开始填写。"
