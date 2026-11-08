@echo off
setlocal EnableDelayedExpansion
chcp 65001 >nul
cd /d "%~dp0"

echo ============================================================================
echo   CAI DAT MOI TRUONG - CONG CU CHI HO LUONG NGAN HANG (ChiHoLuongTool V1)
echo ============================================================================
echo.

where python >nul 2>nul
if errorlevel 1 (
    echo [LOI] Khong tim thay Python tren may tinh.
    echo Vui long cai dat Python 3.10 tro len va tick chon "Add Python to PATH".
    pause
    exit /b 1
)

if not exist ".venv\Scripts\python.exe" (
    echo [1/3] Dang tao moi truong ao (.venv)...
    python -m venv .venv
    if errorlevel 1 (
        echo [LOI] Khong the tao moi truong ao .venv.
        pause
        exit /b 1
    )
) else (
    echo [1/3] Moi truong ao (.venv) da ton tai.
)

echo [2/3] Dang nang cap pip va cai dat thu vien tu requirements.txt...
".venv\Scripts\python.exe" -m pip install --upgrade pip
".venv\Scripts\python.exe" -m pip install -r requirements.txt
if errorlevel 1 (
    echo [LOI] Cai dat thu vien that bai. Vui long kiem tra ket noi mang.
    pause
    exit /b 1
)

echo [3/3] Dang khoi tao thu muc du lieu va cau hinh mac dinh...
if not exist "templates" mkdir "templates"
if not exist "data\backup" mkdir "data\backup"
if not exist "data\output" mkdir "data\output"
if not exist "data\reports" mkdir "data\reports"
if not exist "data\logs" mkdir "data\logs"

echo.
echo ============================================================================
echo   HOAN TAT CAI DAT! Ban co the chay chuong trinh bang file run.bat
echo ============================================================================
pause
exit /b 0
