@echo off
chcp 65001 >nul
cd /d "%~dp0"
python -c "import pythoncom; pythoncom.CoInitialize()" 2>nul
python app.py
pause