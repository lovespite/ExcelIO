param([ValidateSet('Debug', 'Release')][string]$Configuration = 'Release', [switch]$SkipTests)
$ErrorActionPreference = 'Stop'
$projectRoot = Split-Path $PSScriptRoot -Parent
Push-Location $projectRoot
try {
    if (-not $SkipTests) {
        dotnet test .\QuickExcel.Core.Test\QuickExcel.Core.Test.csproj -c $Configuration -v minimal
        if ($LASTEXITCODE -ne 0) { throw 'QuickExcel tests failed.' }
    }
    dotnet publish .\QuickExcel.Desktop\QuickExcel.Desktop.csproj -c $Configuration --self-contained false -o .\artifacts\QuickExcel -v minimal
    if ($LASTEXITCODE -ne 0) { throw 'QuickExcel publish failed.' }
    Write-Host "Ready: $projectRoot\artifacts\QuickExcel\QuickExcel.exe"
} finally { Pop-Location }
