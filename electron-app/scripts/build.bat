@echo off
REM Build script for Risk Management Desktop App
REM Run this on Windows to build the Windows .exe

echo ==========================================
echo   Risk Management - Build Script
echo ==========================================

cd /d "%~dp0\.."
set ROOT_DIR=%cd%\..

echo.
echo [1/5] Installing Electron dependencies...
call npm install

echo.
echo [2/5] Building React dashboard...
cd "%ROOT_DIR%\dashboard"
call npm install
call npm run build
mkdir "%ROOT_DIR%\electron-app\build\webapp" 2>nul
xcopy /E /Y dist\* "%ROOT_DIR%\electron-app\build\webapp\"

echo.
echo [3/5] Installing PyInstaller...
pip install pyinstaller

echo.
echo [4/5] Building Python server...
cd "%ROOT_DIR%\electron-app\python"
pyinstaller --clean --noconfirm risk_server.spec

echo.
echo [5/5] Building Electron app...
cd "%ROOT_DIR%\electron-app"
call npm run build:win

echo.
echo ==========================================
echo   Build Complete!
echo ==========================================
echo.
echo Windows installer: electron-app\dist\Risk Management Setup.exe
echo.
pause
