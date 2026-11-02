$ErrorActionPreference = 'Stop'
Push-Location -LiteralPath $PSScriptRoot
try {
    py -3.14 -m unittest discover -s tests -v
    if ($LASTEXITCODE -ne 0) { throw 'Tests failed.' }
    py -3.14 -m PyInstaller --noconfirm HybeFiller.spec
    if ($LASTEXITCODE -ne 0) { throw 'Build failed.' }
    Copy-Item -LiteralPath (Join-Path $PSScriptRoot 'Template.xlsx') -Destination (Join-Path $PSScriptRoot 'dist\HybeFiller_v1.3.1_portable\Template.xlsx')
    Copy-Item -LiteralPath (Join-Path $PSScriptRoot 'USER_GUIDE_KO.txt') -Destination (Join-Path $PSScriptRoot 'dist\HybeFiller_v1.3.1_portable\USER_GUIDE_KO.txt')
} finally {
    Pop-Location
}
