@echo off
echo Building executable...

pip install -r requirements.txt

pyinstaller app.spec --clean

echo.
echo Build complete! Executable is in dist folder.
pause