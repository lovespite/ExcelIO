param(
    [Parameter(Mandatory = $true)]
    [string]$WasmPublishDir,

    [string]$NpmDir = "$PSScriptRoot\..",

    [string]$Version
)

$ErrorActionPreference = "Stop"
$NpmDir = [System.IO.Path]::GetFullPath($NpmDir)
$distDir = "$NpmDir\dist"

Write-Host "Preparing npm package at $distDir"

# ── Clean dist ──
if (Test-Path $distDir) {
    Remove-Item -Recurse -Force $distDir
}
New-Item -ItemType Directory -Force $distDir | Out-Null

# ── Copy _framework ──
$frameworkSource = "$WasmPublishDir\_framework"
if (-not (Test-Path $frameworkSource)) {
    Write-Host "ERROR: _framework not found at $frameworkSource" -ForegroundColor Red
    exit 1
}
Write-Host "  Copying _framework ..."
Copy-Item -Path $frameworkSource -Destination "$distDir\_framework\" -Recurse -Force

# ── Remove compressed variants (*.br, *.gz) to reduce npm package size ──
Write-Host "  Removing pre-compressed variants (.br, .gz) ..."
$removed = 0
Get-ChildItem "$distDir\_framework" -Include "*.br", "*.gz" -Recurse | ForEach-Object {
    Remove-Item $_.FullName -Force
    $removed++
}
Write-Host "  Removed $removed compressed files"

# ── Copy JS and .d.ts source ──
Copy-Item -Path "$NpmDir\src\excelio.js" -Destination "$distDir\excelio.js" -Force
Copy-Item -Path "$NpmDir\src\excelio.d.ts" -Destination "$distDir\excelio.d.ts" -Force

# ── Inject version into package.json ──
if ($Version) {
    Write-Host "  Setting version: $Version"
    $pkgPath = "$NpmDir\package.json"
    $pkg = Get-Content $pkgPath -Raw | ConvertFrom-Json
    $pkg.version = $Version
    $pkg | ConvertTo-Json -Depth 10 | Set-Content $pkgPath -Encoding UTF8
}

# ── Report ──
$fileCount = (Get-ChildItem $distDir -Recurse -File).Count
$totalSize = "{0:N1} MB" -f ((Get-ChildItem $distDir -Recurse -File | Measure-Object -Property Length -Sum).Sum / 1MB)

Write-Host "npm package prepared: $fileCount files, $totalSize" -ForegroundColor Green
