@echo off
chcp 65001 >nul 2>&1
title Excel 转 PDF 工具 - 一键打包 EXE

echo.
echo ========================================
echo    一键打包为独立 EXE 文件
echo ========================================
echo.

REM 检查 Python
python --version >nul 2>&1
if %errorlevel% neq 0 (
    echo ❌ 未检测到 Python，请先安装 Python
    pause
    exit /b 1
)

REM 安装依赖
echo 📥 正在安装打包工具...
pip install pywin32 pyinstaller -q
echo ✅ 依赖安装完成
echo.

REM 开始打包
echo 📦 正在打包，这可能需要 1-3 分钟...
echo.

pyinstaller -F -w --name "Excel转PDF工具" --clean "%~dp0main.py"

if %errorlevel% equ 0 (
    echo.
    echo ========================================
    echo ✅ 打包成功！
    echo.
    echo 📂 EXE 文件位置:
    echo    %~dp0dist\Excel转PDF工具.exe
    echo.
    echo 把这个 EXE 文件发给业务同学即可！
    echo ========================================
    echo.

    REM 自动打开 dist 文件夹
    explorer "%~dp0dist"
) else (
    echo.
    echo ❌ 打包失败，请查看上方错误信息
)

echo.
pause
