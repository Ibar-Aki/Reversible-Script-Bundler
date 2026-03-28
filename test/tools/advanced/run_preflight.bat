@echo off
setlocal
for %%I in ("%~dp0..\..") do set "ROOT_DIR=%%~fI\"
set "PS_EXE=powershell.exe"
where pwsh >nul 2>nul && set "PS_EXE=pwsh"

if "%~1"=="" goto :usage
if /I "%~1"=="-h" goto :usage
if /I "%~1"=="--help" goto :usage
if /I "%~1"=="/?" goto :usage

"%PS_EXE%" -NoLogo -NoProfile -ExecutionPolicy Bypass -File "%ROOT_DIR%scripts\preflight_excel.ps1" %*
exit /b %errorlevel%

:usage
"%PS_EXE%" -NoLogo -NoProfile -ExecutionPolicy Bypass -File "%ROOT_DIR%scripts\show_usage.ps1" -CommandName run_preflight
exit /b 1
