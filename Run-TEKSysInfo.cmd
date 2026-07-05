@echo off
setlocal EnableExtensions

set "SCRIPT_DIR=%~dp0"
set "SCRIPT_PATH=%SCRIPT_DIR%TEKSysInfo.ps1"

if not exist "%SCRIPT_PATH%" (
    echo TEKSysInfo.ps1 was not found next to this launcher.
    echo Expected: "%SCRIPT_PATH%"
    pause
    exit /b 1
)

net session >nul 2>&1
if %errorlevel% neq 0 (
    echo Requesting administrator approval...
    powershell.exe -NoProfile -ExecutionPolicy Bypass -Command "Start-Process -FilePath '%~f0' -Verb RunAs -WorkingDirectory '%SCRIPT_DIR%'"
    exit /b
)

pushd "%SCRIPT_DIR%"
powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%SCRIPT_PATH%"
set "EXIT_CODE=%errorlevel%"
popd

if not "%EXIT_CODE%"=="0" (
    echo.
    echo TEKSysInfo exited with code %EXIT_CODE%.
    pause
)

exit /b %EXIT_CODE%
