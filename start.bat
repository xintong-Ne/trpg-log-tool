@echo off
cd /d "%~dp0"

where python >nul 2>nul
if errorlevel 1 (
  echo 没有找到 Python。请先安装 Python 3。
  pause
  exit /b 1
)

if not exist ".venv" (
  python -m venv .venv
)

call .venv\Scripts\activate.bat
python -m pip install --upgrade pip
python -m pip install -r requirements.txt
python -m streamlit run app.py --browser.gatherUsageStats false
