@echo off
setlocal EnableDelayedExpansion
for %%I in ("%~f0") do set "SCRIPT_DIR=%%~dpI"
set "BUNDLE_ORIGINAL_CMDCMDLINE=!CMDCMDLINE!"
setlocal DisableDelayedExpansion
set "EXIT_PARENT_FLAG=%TEMP%\bundle_exit_parent_%RANDOM%_%RANDOM%.flag"
if exist "%EXIT_PARENT_FLAG%" del "%EXIT_PARENT_FLAG%" >nul 2>nul
powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%SCRIPT_DIR%bundle_launcher.ps1" -DefaultMode Bundle -SystemPath "%SCRIPT_DIR%bundle_system.ps1" -BatchPath "%~f0" -ExitParentFlagPath "%EXIT_PARENT_FLAG%"
set "EXIT_CODE=%ERRORLEVEL%"
if exist "%EXIT_PARENT_FLAG%" (
    del "%EXIT_PARENT_FLAG%" >nul 2>nul
    endlocal & exit %EXIT_CODE%
)
endlocal & exit /b %EXIT_CODE%
