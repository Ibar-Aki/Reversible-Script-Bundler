@echo off
setlocal

set "ROOT_DIR=%~dp0"
set "PS_EXE=powershell.exe"
where pwsh >nul 2>nul && set "PS_EXE=pwsh"

if "%~1"=="" goto :menu
if /I "%~1"=="-h" goto :usage
if /I "%~1"=="--help" goto :usage
if /I "%~1"=="/?" goto :usage

if /I "%~1"=="-RunAll" goto :route_run_all
if /I "%~1"=="-PromptBundle" goto :route_prompt_bundle
if /I "%~1"=="-SelfTest" goto :route_self_test
if /I "%~1"=="-Preflight" goto :usage
if /I "%~1"=="-Extract" goto :usage
if /I "%~1"=="-Pack" goto :usage
if /I "%~1"=="-Verify" goto :usage
if /I "%~1"=="-Rebuild" goto :usage

goto :run_all

:run_all
call "%ROOT_DIR%tools\user\run_all.bat" %*
exit /b %errorlevel%

:route_run_all
call "%ROOT_DIR%tools\user\run_all.bat" %2 %3 %4 %5 %6 %7 %8 %9
exit /b %errorlevel%

:route_prompt_bundle
call "%ROOT_DIR%tools\user\run_prompt_bundle.bat" %2 %3 %4 %5 %6 %7 %8 %9
exit /b %errorlevel%

:route_self_test
call "%ROOT_DIR%tools\user\run_self_test.bat" %2 %3 %4 %5 %6 %7 %8 %9
exit /b %errorlevel%

:menu
"%PS_EXE%" -NoLogo -NoProfile -ExecutionPolicy Bypass -File "%ROOT_DIR%scripts\launch_menu.ps1" -ProjectRoot "%ROOT_DIR%"
exit /b %errorlevel%

:usage
"%PS_EXE%" -NoLogo -NoProfile -ExecutionPolicy Bypass -File "%ROOT_DIR%scripts\show_usage.ps1" -CommandName Excel2LLM
exit /b 1
