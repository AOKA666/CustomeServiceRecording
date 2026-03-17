@echo off
chcp 65001 >nul
title 客服记录台

echo 正在检查 Python...
python --version >nul 2>&1
if errorlevel 1 (
    echo [错误] 未找到 Python，请先安装 Python 3.8+
    pause
    exit /b
)

echo 正在安装依赖...
pip install flask requests -q

echo.
echo ============================================
echo   客服记录台 启动中...
echo   请在浏览器访问: http://127.0.0.1:5000
echo ============================================
echo.

start "" http://127.0.0.1:5000
python app.py

pause
