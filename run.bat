@echo off
setlocal
chcp 65001 >nul
set "PYTHONUTF8=1"
set "PYTHONIOENCODING=utf-8"
cd /d "%~dp0"

if exist ".venv\Scripts\python.exe" (
    ".venv\Scripts\python.exe" app.py %*
) else (
    python app.py %*
)

if errorlevel 1 (
    echo.
    echo [LOI] Chuong trinh ket thuc voi ma loi %errorlevel%.
    echo Xem chi tiet tai thu muc data\logs\
    pause
    exit /b %errorlevel%
)
exit /b 0
