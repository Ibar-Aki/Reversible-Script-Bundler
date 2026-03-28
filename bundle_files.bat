@echo off
setlocal
set "SHOULD_PAUSE=1"

:parse_args
if "%~1"=="" goto run
if /I "%~1"=="--no-pause" set "SHOULD_PAUSE=0"
shift
goto parse_args

:run
powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%~dp0bundle_system.ps1" -Mode Bundle
set "EXIT_CODE=%ERRORLEVEL%"
if "%SHOULD_PAUSE%"=="1" pause
endlocal & exit /b %EXIT_CODE%
