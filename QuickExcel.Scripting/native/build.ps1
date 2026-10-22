param([switch]$Rebuild)
$ErrorActionPreference = 'Stop'
# Normalize PATH casing before invoking the .NET Framework MSBuild toolchain.
$quickExcelPathValue = $env:Path
Remove-Item Env:PATH -ErrorAction SilentlyContinue
Remove-Item Env:Path -ErrorAction SilentlyContinue
$env:Path = $quickExcelPathValue
$nativeRoot = $PSScriptRoot
$output = Join-Path $nativeRoot 'bin\quickexcel-js.dll'
if ((Test-Path -LiteralPath $output) -and -not $Rebuild) { return }
$version = '0.17.0'
$expectedHash = '9DE1BF573606D2AEBBBAB6217E616A66E09E3613518489B37B9B3157CF398DCC'
$vendor = Join-Path $nativeRoot 'vendor'
$archive = Join-Path $vendor "quickjs-ng-v$version.zip"
$sourceRoot = Join-Path $vendor "quickjs-$version"
New-Item -ItemType Directory -Force -Path $vendor | Out-Null
if (-not (Test-Path -LiteralPath $archive)) {
    Invoke-WebRequest "https://github.com/quickjs-ng/quickjs/archive/refs/tags/v$version.zip" -OutFile $archive
}
if ((Get-FileHash -LiteralPath $archive -Algorithm SHA256).Hash -ne $expectedHash) {
    throw 'QuickJS source archive SHA-256 does not match the pinned source.'
}
if (-not (Test-Path -LiteralPath (Join-Path $sourceRoot 'quickjs.h'))) {
    Expand-Archive -LiteralPath $archive -DestinationPath $vendor
}
$cmake = (Get-Command cmake -ErrorAction SilentlyContinue).Source
if (-not $cmake) { throw 'Install CMake 3.21+ and Visual Studio C++ build tools to build the native runtime.' }
$vswhere = Join-Path ${env:ProgramFiles(x86)} 'Microsoft Visual Studio\Installer\vswhere.exe'
if (-not (Test-Path -LiteralPath $vswhere)) { throw 'Visual Studio C++ build tools are required.' }
$installation = & $vswhere -latest -products * -requires Microsoft.VisualStudio.Component.VC.Tools.x86.x64 -property installationPath
if (-not $installation) { throw 'No installed Visual Studio C++ compiler was found.' }
$vsVersion = & $vswhere -latest -products * -requires Microsoft.VisualStudio.Component.VC.Tools.x86.x64 -property installationVersion
$major = [int]($vsVersion.Split('.')[0])
$generator = if ($major -ge 18) { 'Visual Studio 18 2026' } else { 'Visual Studio 17 2022' }
& $cmake -S $nativeRoot -B (Join-Path $nativeRoot 'build') -G $generator -A x64 "-DCMAKE_GENERATOR_INSTANCE=$installation"
if ($LASTEXITCODE -ne 0) { throw 'CMake configuration failed.' }
& $cmake --build (Join-Path $nativeRoot 'build') --config Release --target quickexcel-js --parallel
if ($LASTEXITCODE -ne 0) { throw 'QuickJS native build failed.' }
New-Item -ItemType Directory -Force -Path (Join-Path $nativeRoot 'bin') | Out-Null
Copy-Item -LiteralPath (Join-Path $nativeRoot 'build\Release\quickexcel-js.dll') -Destination $output -Force
Write-Host "QuickJS-NG $version runtime: $output"
