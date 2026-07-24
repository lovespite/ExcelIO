param(
    [ValidateSet("Debug", "Release")]
    [string]$Configuration = "Release",

    [string]$OutputDir = "$PSScriptRoot\..\artifacts",

    [string]$Version,

    [switch]$SkipTests
)

$ErrorActionPreference = "Stop"
$root = (Resolve-Path "$PSScriptRoot\..").Path
$scriptsDir = "$PSScriptRoot"
$OutputDir = [System.IO.Path]::GetFullPath($OutputDir)

# ── Run build + tests first ──

Write-Host "=== ExcelIO Pack Script ===" -ForegroundColor Cyan
Write-Host "Configuration: $Configuration"
Write-Host "Output: $OutputDir"
Write-Host ""

$buildResult = & "$scriptsDir\build.ps1" -Configuration $Configuration -SkipTests:$SkipTests
if ($LASTEXITCODE -ne 0) {
    Write-Host "Build failed. Aborting pack." -ForegroundColor Red
    exit 1
}

# ── Apply version override ──

if ($Version) {
    Write-Host "Overriding version to $Version ..." -ForegroundColor Yellow
    $csprojFiles = @(
        "$root\ExcelIO\ExcelIO.csproj",
        "$root\ExcelIO.NetStandard\ExcelIO.NetStandard.csproj"
    )
    $tempFiles = @{}
    foreach ($file in $csprojFiles) {
        $bak = "$file.version.bak"
        Copy-Item $file $bak
        $tempFiles[$file] = $bak
        $xml = [xml](Get-Content $file)
        $versionNode = $xml.Project.PropertyGroup.Version
        if ($versionNode) {
            $versionNode | ForEach-Object { $_.InnerText = $Version }
        } else {
            $pg = $xml.Project.PropertyGroup[0]
            $vn = $xml.CreateElement("Version")
            $vn.InnerText = $Version
            $pg.AppendChild($vn) | Out-Null
        }
        $xml.Save($file)
    }
}

try {
    # ── Clean output ──

    if (Test-Path $OutputDir) {
        Remove-Item -Recurse -Force $OutputDir
    }
    New-Item -ItemType Directory -Force $OutputDir | Out-Null

    # ── Pack NuGet packages ──

    Write-Host ""
    Write-Host "── Packing NuGet ──" -ForegroundColor Cyan

    $packProjects = @(
        @{ Path = "$root\ExcelIO\ExcelIO.csproj";                  Name = "ExcelIO" },
        @{ Path = "$root\ExcelIO.NetStandard\ExcelIO.NetStandard.csproj"; Name = "ExcelIO.NetStandard" }
    )

    foreach ($proj in $packProjects) {
        Write-Host "  $($proj.Name) ... " -NoNewline
        $result = dotnet pack $proj.Path -c $Configuration -o $OutputDir -v minimal 2>&1
        if ($LASTEXITCODE -eq 0) {
            Write-Host "OK" -ForegroundColor Green
        } else {
            Write-Host "FAILED" -ForegroundColor Red
            Write-Host ($result -join "`n")
            exit 1
        }
    }

    # ── Publish WASM ──

    Write-Host ""
    Write-Host "── Publishing ExcelIO.Web ──" -ForegroundColor Cyan

    $wasmOut = "$OutputDir\publish\wwwroot"
    Write-Host "  Publishing to $wasmOut ... " -NoNewline
    $result = dotnet publish "$root\ExcelIO.Web\ExcelIO.Web.csproj" -c $Configuration -o $wasmOut -v minimal 2>&1
    if ($LASTEXITCODE -eq 0) {
        Write-Host "OK" -ForegroundColor Green
    } else {
        Write-Host "FAILED" -ForegroundColor Red
        Write-Host ($result -join "`n")
        exit 1
    }

    # ── Prepare npm package ──

    Write-Host ""
    Write-Host "── Preparing npm package ──" -ForegroundColor Cyan

    $npmPrepareScript = "$root\ExcelIO.Web\npm\scripts\prepare-npm.ps1"
    if (Test-Path $npmPrepareScript) {
        $wasmWwwroot = "$wasmOut\wwwroot"
        if (-not (Test-Path "$wasmWwwroot\_framework")) {
            Write-Host "  WARNING: _framework not found at $wasmWwwroot, trying $wasmOut"
            $wasmWwwroot = $wasmOut
        }
        Write-Host "  WASM source: $wasmWwwroot"
        & $npmPrepareScript -WasmPublishDir $wasmWwwroot -Version $Version
        if ($LASTEXITCODE -ne 0) {
            Write-Host "npm prepare failed!" -ForegroundColor Red
            exit 1
        }

        # Optionally run npm pack to verify
        $npmDir = "$root\ExcelIO.Web\npm"
        Write-Host "  Running npm pack ... " -NoNewline
        Push-Location $npmDir
        $packResult = npm pack --pack-destination $OutputDir 2>&1
        Pop-Location
        if ($LASTEXITCODE -eq 0) {
            Write-Host "OK" -ForegroundColor Green
        } else {
            Write-Host "SKIPPED (npm not available or pack failed)" -ForegroundColor Yellow
        }
    }

    # ── Summary ──

    Write-Host ""
    Write-Host "Artifacts:" -ForegroundColor Cyan
    Get-ChildItem $OutputDir -Recurse -File | ForEach-Object {
        $size = "{0:N1} KB" -f ($_.Length / 1KB)
        Write-Host "  $($_.FullName.Replace($OutputDir, '').TrimStart('\'))  ($size)"
    }

    Write-Host ""
    Write-Host "Pack complete." -ForegroundColor Green

} finally {
    # ── Restore original version files ──

    if ($Version) {
        foreach ($kv in $tempFiles.GetEnumerator()) {
            $bak = $kv.Value
            if (Test-Path $bak) {
                Move-Item $bak $kv.Key -Force
            }
        }
    }
}
