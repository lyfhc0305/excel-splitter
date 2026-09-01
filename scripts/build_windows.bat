@echo off
setlocal

REM cd to project root (this script lives in scripts\)
cd /d "%~dp0.."

if not exist .venv (
    py -m venv .venv
)

call .venv\Scripts\activate.bat
python -m pip install --upgrade pip
python -m pip install -r requirements.txt

if exist build rmdir /s /q build
if exist dist rmdir /s /q dist

pyinstaller ^
  --noconfirm ^
  --clean ^
  --windowed ^
  --name ExcelSplitter-Windows ^
  src\excel_splitter.py

echo.
echo Build complete. Output: dist\ExcelSplitter-Windows
endlocal
