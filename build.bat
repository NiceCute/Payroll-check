@echo off
setlocal EnableDelayedExpansion
chcp 65001 >nul
set "PYTHONUTF8=1"
set "PYTHONIOENCODING=utf-8"
cd /d "%~dp0"

echo ============================================================================
echo   DONG GOI BAN PHAT HANH (.EXE) - ChiHoLuongTool V1
echo ============================================================================
echo.

set "PY_CMD=python"
if exist ".venv\Scripts\python.exe" (
    set "PY_CMD=.venv\Scripts\python.exe"
)

echo [1/4] Kiem tra PyInstaller...
"%PY_CMD%" -m PyInstaller --version >nul 2>nul
if errorlevel 1 (
    echo PyInstaller chua duoc cai dat. Dang cai dat PyInstaller...
    "%PY_CMD%" -m pip install pyinstaller
    if errorlevel 1 (
        echo [LOI] Khong the cai dat PyInstaller.
        pause
        exit /b 1
    )
)

echo [2/4] Dang chay PyInstaller voi pyinstaller.spec...
"%PY_CMD%" -m PyInstaller --noconfirm --clean pyinstaller.spec
if errorlevel 1 (
    echo [LOI] Dong goi PyInstaller that bai!
    pause
    exit /b 1
)

echo [3/4] Sao chep templates va tao cau truc thu muc data cho ban phat hanh...
if not exist "dist\ChiHoLuongTool\templates" mkdir "dist\ChiHoLuongTool\templates"
copy /y "templates\*.json" "dist\ChiHoLuongTool\templates\" >nul
if not exist "dist\ChiHoLuongTool\data\backup" mkdir "dist\ChiHoLuongTool\data\backup"
if not exist "dist\ChiHoLuongTool\data\output" mkdir "dist\ChiHoLuongTool\data\output"
if not exist "dist\ChiHoLuongTool\data\reports" mkdir "dist\ChiHoLuongTool\data\reports"
if not exist "dist\ChiHoLuongTool\data\logs" mkdir "dist\ChiHoLuongTool\data\logs"
copy /y "README.md" "dist\ChiHoLuongTool\" >nul

echo [4/4] Hoan tat dong goi!
echo Ban phat hanh nam tai: dist\ChiHoLuongTool\ChiHoLuongTool.exe
echo ============================================================================
if /i not "%~1"=="--nopause" pause
exit /b 0
