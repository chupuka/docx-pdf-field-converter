@echo off
echo Installing dependencies...
pip install -r requirements.txt

echo.
echo Building executable with PyInstaller...
pyinstaller app.spec --clean --noconfirm

echo.
echo ========================================
echo Build complete! 
echo Executable: dist\DocxToPdfConverter.exe
echo ========================================
pause