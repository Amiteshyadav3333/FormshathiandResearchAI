@echo off
setlocal
cd /d %~dp0
if not exist .venv\Scripts\python.exe py -3 -m venv .venv
.venv\Scripts\python.exe -m pip install --upgrade pip
.venv\Scripts\python.exe -m pip install -r requirements-desktop.txt
.venv\Scripts\pyinstaller.exe --noconfirm --clean --windowed --name FormSaathi --add-data "index.html;." --add-data "app.js;." --add-data "styles.css;." --collect-all webview --collect-all pdfplumber desktop.py
echo Windows app ready: dist\FormSaathi\FormSaathi.exe
pause
