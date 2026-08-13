@echo off
setlocal
cd /d "%~dp0"
if exist "%~dp0Start.vbs" (
    wscript.exe "%~dp0Start.vbs"
    exit /b
)
powershell.exe -NoProfile -ExecutionPolicy Bypass -STA -WindowStyle Hidden -File "%~dp0app\Start-Portable.ps1"
