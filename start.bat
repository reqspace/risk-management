@echo off
REM Risk Management - Windows Start Script
REM Double-click this file to start the application

cd /d "%~dp0"

echo ==========================================
echo    Risk Management Dashboard
echo ==========================================
echo.

REM Check if .env exists
if not exist .env (
    echo ERROR: .env file not found!
    echo Please copy .env.example to .env and fill in your settings.
    echo.
    pause
    exit /b 1
)

REM Check if node_modules exists
if not exist "dashboard\node_modules" (
    echo Installing dashboard dependencies (first run only)...
    cd dashboard
    call npm install
    cd ..
    echo.
)

echo Starting Flask API server...
start /b python server.py

timeout /t 2 /nobreak >nul

echo Starting Dashboard...
cd dashboard
start /b npm run dev
cd ..

timeout /t 3 /nobreak >nul

echo.
echo ==========================================
echo    Application Started!
echo ==========================================
echo.
echo Dashboard: http://localhost:3000
echo API Server: http://localhost:5001
echo.
echo Close this window to stop all services
echo.

REM Open browser
start http://localhost:3000

REM Keep window open
pause
