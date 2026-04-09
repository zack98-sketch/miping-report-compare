@echo off
chcp 65001 >nul 2>&1
title 密评报告对比工具
echo.
echo   ╔══════════════════════════════════════╗
echo   ║    📊 密评报告附录D 对比工具 v1.0     ║
echo   ║    上传两份报告，即时生成HTML对比      ║
echo   ╚══════════════════════════════════════╝
echo.

cd /d "%~dp0"

:: Check Python
where python >nul 2>&1
if %errorlevel% neq 0 (
    echo [错误] 未找到 Python。请先安装 Python 3.8+
    pause & exit /b 1
)

:: Install dependencies silently
echo [1/2] 检查依赖...
pip show flask >nul 2>&1
if %errorlevel% neq 0 (
    echo       正在安装 flask...
    pip install flask -q
)

:: Start server
echo [2/2] 启动服务器...
echo.
echo   浏览器即将自动打开: http://127.0.0.1:5678
echo   如未自动打开，请手动访问上方地址
echo.
echo   按 Ctrl+C 停止服务
echo.

python app.py

pause
