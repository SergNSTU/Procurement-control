@echo off
setlocal
cd /d "%~dp0"

if not exist "%~dp0Procurement_control_portable\Start.bat" (
    echo Не найдена папка Procurement_control_portable.
    echo Скачайте и распакуйте репозиторий целиком, затем повторите запуск.
    pause
    exit /b 1
)

call "%~dp0Procurement_control_portable\Start.bat"
exit /b %ERRORLEVEL%
