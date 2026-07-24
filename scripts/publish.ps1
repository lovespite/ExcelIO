param(
    [Parameter(Mandatory = $true)]
    [string]$ApiKey,

    [string]$Source = "https://api.nuget.org/v3/index.json",

    [string]$PackageDir = "$PSScriptRoot\..\artifacts"
)

$ErrorActionPreference = "Stop"
$PackageDir = [System.IO.Path]::GetFullPath($PackageDir)

Write-Host "=== ExcelIO Publish Script ===" -ForegroundColor Cyan
Write-Host "Source: $Source"
Write-Host "PackageDir: $PackageDir"
Write-Host ""

if (-not (Test-Path $PackageDir)) {
    Write-Host "Package directory not found: $PackageDir" -ForegroundColor Red
    Write-Host "Run pack.ps1 first." -ForegroundColor Yellow
    exit 1
}

$nupkgs = Get-ChildItem $PackageDir -Filter "*.nupkg" | Sort-Object Name
if ($nupkgs.Count -eq 0) {
    Write-Host "No .nupkg files found in $PackageDir" -ForegroundColor Red
    exit 1
}

Write-Host "Packages to push:" -ForegroundColor Cyan
foreach ($pkg in $nupkgs) {
    Write-Host "  $($pkg.Name)"
}
Write-Host ""

$confirm = Read-Host "Push these packages to $Source? (y/N)"
if ($confirm -ne 'y' -and $confirm -ne 'Y') {
    Write-Host "Aborted." -ForegroundColor Yellow
    exit 0
}

Write-Host ""

foreach ($pkg in $nupkgs) {
    Write-Host "  Pushing $($pkg.Name) ... " -NoNewline
    $result = dotnet nuget push $pkg.FullName --api-key $ApiKey --source $Source --skip-duplicate 2>&1
    if ($LASTEXITCODE -eq 0) {
        Write-Host "OK" -ForegroundColor Green
    } else {
        Write-Host "FAILED" -ForegroundColor Red
        Write-Host ($result -join "`n")
        exit 1
    }
}

Write-Host ""
Write-Host "Publish complete." -ForegroundColor Green
