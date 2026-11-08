@echo off
chcp 65001 >nul
cd /d "%~dp0"
REM Chạy toàn bộ kiểm thử tự động.
REM Muốn kiểm thử thêm với file lương thật:  set CHL_REAL_SOURCE=C:\duong\dan\file_luong.xlsx
set PYTHONIOENCODING=utf-8
if exist ".venv\Scripts\python.exe" (set PY=.venv\Scripts\python.exe) else (set PY=python)
%PY% -m unittest discover -s tests -t . -v
pause
