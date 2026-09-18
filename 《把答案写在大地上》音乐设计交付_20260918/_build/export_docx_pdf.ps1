$ErrorActionPreference = 'Stop'
$docxPath = (Get-ChildItem -LiteralPath (Split-Path $PSScriptRoot -Parent) -File -Filter '*.docx' | Select-Object -First 1).FullName
$renderDir = Join-Path $PSScriptRoot 'docx_render'
New-Item -ItemType Directory -Force -Path $renderDir | Out-Null
$qaDocxPath = Join-Path $renderDir 'music_design_qa.docx'
$pdfPath = Join-Path $renderDir 'music_design_qa.pdf'
Copy-Item -LiteralPath $docxPath -Destination $qaDocxPath -Force
$word = New-Object -ComObject Word.Application
$word.Visible = $false
$word.DisplayAlerts = 0
try {
    $doc = $word.Documents.Open($qaDocxPath)
    $doc.ExportAsFixedFormat($pdfPath, 17)
    $doc.Close($false)
}
finally {
    $word.Quit()
}
Write-Output $pdfPath
