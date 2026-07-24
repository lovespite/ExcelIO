param(
    [ValidateSet("Debug", "Release")]
    [string]$Configuration = "Release",

    [switch]$SkipTests
)

$ErrorActionPreference = "Stop"
$root = (Resolve-Path "$PSScriptRoot\..").Path

Write-Host "=== ExcelIO Build Script ===" -ForegroundColor Cyan
Write-Host "Configuration: $Configuration"
Write-Host ""

$projects = @(
    @{ Path = "$root\ExcelIO\ExcelIO.csproj";                  Name = "ExcelIO (net10.0)" },
    @{ Path = "$root\ExcelIO.NetStandard\ExcelIO.NetStandard.csproj"; Name = "ExcelIO.NetStandard (netstandard2.0)" },
    @{ Path = "$root\ExcelIO.Web\ExcelIO.Web.csproj";          Name = "ExcelIO.Web (net8.0)" }
)

$testProjects = @(
    @{ Path = "$root\ExcelIO.Test\ExcelIO.Test.csproj";                       Name = "ExcelIO.Test" },
    @{ Path = "$root\ExcelIO.NetStandard.Test\ExcelIO.NetStandard.Test.csproj"; Name = "ExcelIO.NetStandard.Test" }
)

$failed = @()

# ── Build ──

Write-Host "── Building ──" -ForegroundColor Cyan

foreach ($proj in $projects) {
    Write-Host "  $($proj.Name) ... " -NoNewline
    $result = dotnet build $proj.Path -c $Configuration -v minimal 2>&1
    if ($LASTEXITCODE -eq 0) {
        Write-Host "OK" -ForegroundColor Green
    } else {
        Write-Host "FAILED" -ForegroundColor Red
        Write-Host ($result -join "`n")
        $failed += $proj.Name
    }
}

# ── Test ──

if (-not $SkipTests) {
    Write-Host ""
    Write-Host "── Testing ──" -ForegroundColor Cyan

    foreach ($test in $testProjects) {
        Write-Host "  $($test.Name) ... " -NoNewline
        $result = dotnet test $test.Path -c $Configuration -v minimal 2>&1
        if ($LASTEXITCODE -eq 0) {
            Write-Host "PASSED" -ForegroundColor Green
        } else {
            Write-Host "FAILED" -ForegroundColor Red
            Write-Host ($result -join "`n")
            $failed += $test.Name
        }
    }
}

# ── Summary ──

Write-Host ""
if ($failed.Count -eq 0) {
    Write-Host "All projects built and tests passed." -ForegroundColor Green
    exit 0
} else {
    Write-Host "Failures: $($failed -join ', ')" -ForegroundColor Red
    exit 1
}
