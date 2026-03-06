@echo off
chcp 65001 >nul 2>&1
title Excel 转 PDF 工具 - 一键安装并运行

echo.
echo ========================================
echo    Excel 转 PDF 批量转换工具 v1.0.0
echo ========================================
echo.

REM 检查 Python 是否已安装
python --version >nul 2>&1
if %errorlevel% neq 0 (
    echo ❌ 未检测到 Python！
    echo.
    echo 请先安装 Python:
    echo 1. 打开浏览器访问 https://www.python.org/downloads/
    echo 2. 下载并安装 Python 3.10+
    echo 3. 安装时请勾选 "Add Python to PATH" ！
    echo.
    pause
    exit /b 1
)

echo ✅ 检测到 Python:
python --version
echo.

REM 检查并安装 pywin32
echo 正在检查依赖...
python -c "import win32com" >nul 2>&1
if %errorlevel% neq 0 (
    echo 📥 正在安装 pywin32，请稍候...
    pip install pywin32 -q
    if %errorlevel% neq 0 (
        echo ❌ 安装 pywin32 失败，请检查网络连接
        pause
        exit /b 1
    )
    echo ✅ pywin32 安装成功
) else (
    echo ✅ pywin32 已安装
)

echo.
echo 🚀 正在启动程序...
echo.

REM 运行主程序
python "%~dp0main.py"

if %errorlevel% neq 0 (
    echo.
    echo ❌ 程序运行出错，请查看上方错误信息
    pause
)
