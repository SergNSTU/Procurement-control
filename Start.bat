@echo off
setlocal
cd /d "%~dp0"

if not exist "%~dp0app\Start-Portable.ps1" (
    echo Не найдена папка app.
    echo Скачайте и распакуйте репозиторий целиком, затем повторите запуск.
    pause
    exit /b 1
)

where powershell.exe >nul 2>&1
if errorlevel 1 (
    echo Windows PowerShell 5.1 не найден.
    echo Установите Windows PowerShell 5.1 и повторите запуск.
    pause
    exit /b 1
)

echo Запуск Procurement Control...
powershell.exe -NoProfile -ExecutionPolicy Bypass -STA -File "%~dp0app\Start-Portable.ps1"
set "exitCode=%ERRORLEVEL%"

if not "%exitCode%"=="0" (
    echo.
    echo Приложение не запустилось. Подробности сохранены в папке:
    echo %~dp0data\logs
    echo Пришлите файл startup_error_*.txt разработчику.
    pause
)

exit /b %exitCode%
