@echo off
setlocal
for %%I in ("%~dp0..\..") do set "ROOT_DIR=%%~fI\"
set "PS_EXE=powershell.exe"
where pwsh >nul 2>nul && set "PS_EXE=pwsh"
"%PS_EXE%" -NoLogo -NoProfile -ExecutionPolicy Bypass -File "%ROOT_DIR%scripts\run_self_test.ps1" %*
