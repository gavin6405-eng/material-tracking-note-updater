@echo off
chcp 65001 >nul
cd /d "%~dp0"

where python >nul 2>nul
if errorlevel 1 (
  echo 找不到 Python，請先安裝 Python 3.11 或更新版本。
  echo 安裝時請勾選 Add Python to PATH。
  pause
  exit /b 1
)

if not exist ".venv\Scripts\python.exe" (
  echo 第一次執行：正在建立專用環境...
  python -m venv .venv
)

call ".venv\Scripts\activate.bat"
python -m pip install --upgrade pip
python -m pip install -r requirements.txt

echo 正在啟動，瀏覽器會自動開啟...
python -m streamlit run app.py
pause
