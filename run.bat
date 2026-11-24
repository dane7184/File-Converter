@echo off
rem Double-click to start file-converter in your browser.
cd /d "%~dp0"
if not exist ".venv\Scripts\python.exe" (
    echo First run: setting up...
    py -3 -m venv .venv || python -m venv .venv
    ".venv\Scripts\python.exe" -m pip install --disable-pip-version-check -r requirements.txt
)
".venv\Scripts\python.exe" app.py
pause
