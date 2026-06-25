@echo off
echo ============================================
echo ARG Converter - Windows Compilation Script
echo ============================================
echo.

echo [1/3] Installing required packages...
pip install -r requirements.txt
if %errorlevel% neq 0 (
    echo ERROR: Failed to install requirements
    pause
    exit /b 1
)

echo.
echo [2/3] Compiling with PyInstaller...
pyinstaller --noconfirm --onedir --windowed --name "ARG_Converter" app.py
if %errorlevel% neq 0 (
    echo ERROR: Failed to compile application
    pause
    exit /b 1
)

echo.
echo [3/3] Compilation complete!
echo.
echo The executable is located in:
echo   dist\ARG_Converter\ARG_Converter.exe
echo.
echo To run the application:
echo   dist\ARG_Converter\ARG_Converter.exe
echo.
pause
