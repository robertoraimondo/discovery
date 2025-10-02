@echo off
title System Information GUI - Master Version Launcher
echo.
echo ===============================================
echo    SYSTEM INFORMATION GUI - MASTER VERSION
echo ===============================================
echo.
echo Features:
echo  - Auto-elevation to Administrator
echo  - Enhanced button panel (120px height)
echo  - Comprehensive system information
echo  - Services monitoring
echo  - Network shares analysis
echo  - Package management capabilities
echo  - Progress tracking with status updates
echo.
echo Starting Master GUI with Administrator privileges...
echo.

REM Change to script directory
cd /d "%~dp0"

REM Launch PowerShell as Administrator with the Master GUI script
powershell.exe -Command "Start-Process powershell.exe -Verb RunAs -ArgumentList '-ExecutionPolicy Bypass -NoExit -Command \"cd \`\"%~dp0\`\"; .\SystemInfoGUI-Master.ps1\"'"

echo.
echo Master GUI launch command sent to Administrator session.
echo Check for UAC prompt and new Administrator window.
echo.
pause