@echo off
chcp 65001 >nul
cd /d "%~dp0"

where python >nul 2>nul
if errorlevel 1 (
  echo Python not found. Please install Python 3.11 or newer.
  echo Select Add Python to PATH during installation.
  pause
  exit /b 1
)

if not exist ".venv\Scripts\python.exe" (
  echo First run: creating virtual environment...
  python -m venv .venv
)

call ".venv\Scripts\activate.bat"
python -m pip install --upgrade pip
python -m pip install -r requirements.txt

echo Starting material replenishment app...
python -m streamlit run app.py
pause
