@echo off
chcp 65001 >nul 2>&1
title Excel to PDF - Build EXE

echo.
echo ========================================
echo    Build standalone EXE file
echo ========================================
echo.

REM === Find Python ===
set PYTHON_CMD=python
python --version >nul 2>&1
if %errorlevel% neq 0 (
    set PYTHON_CMD=py
    py --version >nul 2>&1
    if %errorlevel% neq 0 (
        echo [FAIL] Python not found!
        echo Please run "Yi Jian Yun Xing.bat" first to install Python.
        echo.
        pause
        exit /b 1
    )
)

echo [OK] Python found:
%PYTHON_CMD% --version
echo.

REM === Install build dependencies ===
echo Installing build tools...
%PYTHON_CMD% -m pip install pywin32 pypdf pyinstaller -q
echo [OK] Build tools ready
echo.

REM === Build EXE ===
echo Building EXE file... (1-3 minutes)
echo.

%PYTHON_CMD% -m PyInstaller -F -w --name "ExcelToPDF" --clean --hidden-import pypdf "%~dp0main.py"

if %errorlevel% equ 0 (
    echo.
    echo ========================================
    echo [OK] Build successful!
    echo.
    echo EXE location:
    echo   %~dp0dist\ExcelToPDF.exe
    echo.
    echo Send this EXE file to your colleagues!
    echo ========================================
    echo.

    REM Open dist folder
    explorer "%~dp0dist"
) else (
    echo.
    echo [FAIL] Build failed. See error above.
)

echo.
pause
