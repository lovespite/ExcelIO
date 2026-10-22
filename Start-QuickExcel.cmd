@echo off
setlocal
cd /d "%~dp0"
if exist "artifacts\QuickExcel\QuickExcel.exe" (
  start "" "artifacts\QuickExcel\QuickExcel.exe" %*
  exit /b 0
)
dotnet run --project "QuickExcel.Desktop\QuickExcel.Desktop.csproj" -- %*
if errorlevel 1 pause
