@echo off
chcp 65001 >nul
title Customer Service Recorder

echo Checking Python...
python --version >nul 2>&1
if errorlevel 1 (
    echo [ERROR] Python not found. Please install Python 3.8+ first.
    pause
    exit /b
)

echo Installing dependencies...
pip install flask requests openpyxl -q

echo.
echo ============================================
echo   Customer service recorder starting...
echo   Local: http://127.0.0.1:5000
echo   LAN:   http://YOUR-LAN-IP:5000
echo   Run ipconfig to find your LAN IPv4 address.
echo ============================================
echo.

start "" http://127.0.0.1:5000
python app.py

pause
