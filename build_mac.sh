#!/bin/sh
set -eu
cd "$(dirname "$0")"
PYTHON="${PYTHON:-./livekit-env/bin/python}"
export PYINSTALLER_CONFIG_DIR="${PYINSTALLER_CONFIG_DIR:-$(pwd)/.pyinstaller-cache}"
"$PYTHON" -m PyInstaller --noconfirm --clean --windowed --name FormSaathi \
  --osx-bundle-identifier com.formsaathi.desktop --add-data "index.html:." \
  --add-data "app.js:." --add-data "styles.css:." \
  --collect-all webview --collect-all pdfplumber --hidden-import=webview.platforms.cocoa desktop.py
echo "Mac app ready: dist/FormSaathi.app"
