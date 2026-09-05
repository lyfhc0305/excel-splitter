@echo off
setlocal

REM cd to project root (this script lives in scripts\)
cd /d "%~dp0.."

if not exist .venv (
    py -m venv .venv
    if errorlevel 1 exit /b 1
)

call .venv\Scripts\activate.bat
if errorlevel 1 exit /b 1
python -m pip install --upgrade pip
if errorlevel 1 exit /b 1
python -m pip install -r requirements.txt
if errorlevel 1 exit /b 1

python -m unittest discover -s tests
if errorlevel 1 exit /b 1

python -m PyInstaller ^
  --noconfirm ^
  --clean ^
  --windowed ^
  --name ExcelSplitter-Windows ^
  src\excel_splitter.py
if errorlevel 1 exit /b 1

echo.
echo Build complete. Output: dist\ExcelSplitter-Windows
endlocal
