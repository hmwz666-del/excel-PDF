@echo off
chcp 65001 >nul 2>&1
title PDF Blank Page Diagnosis

echo.
echo ========================================
echo    PDF Blank Page Diagnosis Tool
echo ========================================
echo.

REM Find Python
python --version >nul 2>&1
if %errorlevel% equ 0 (
    set PYTHON_CMD=python
    goto :run
)
py --version >nul 2>&1
if %errorlevel% equ 0 (
    set PYTHON_CMD=py
    goto :run
)
echo [FAIL] Python not found!
pause
exit /b 1

:run
%PYTHON_CMD% -c "import pypdf" >nul 2>&1
if %errorlevel% neq 0 (
    echo Installing pypdf...
    %PYTHON_CMD% -m pip install pypdf -q
)

%PYTHON_CMD% "%~dp0diagnose.py"
pause
