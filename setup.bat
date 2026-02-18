@echo off
REM CloudAdmin365 Setup Launcher
REM This batch file runs the PowerShell setup script with proper error handling

setlocal enabledelayedexpansion

cd /d "%~dp0"

REM Check if PowerShell is available
where powershell.exe >nul 2>&1
if errorlevel 1 (
    echo Error: PowerShell is not available on this system
    pause
    exit /b 1
)

REM Run the setup script
powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%~dp0setup.ps1"
exit /b %ERRORLEVEL%
