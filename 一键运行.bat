@echo off
chcp 65001 >nul 2>&1
title Excel to PDF Tool - Auto Setup

echo.
echo ========================================
echo    Excel to PDF - Auto Install Script
echo    Excel zhuan PDF - Yi Jian An Zhuang
echo ========================================
echo.

REM === Check if Python exists ===
python --version >nul 2>&1
if %errorlevel% equ 0 (
    echo [OK] Python found:
    python --version
    goto :install_deps
)

py --version >nul 2>&1
if %errorlevel% equ 0 (
    echo [OK] Python found:
    py --version
    goto :install_deps_py
)

echo [!!] Python not found! Auto downloading...
echo.

REM === Download Python installer ===
set PYTHON_URL=https://www.python.org/ftp/python/3.11.9/python-3.11.9-amd64.exe
set INSTALLER=%TEMP%\python-installer.exe

echo Downloading Python 3.11.9 ...
echo Please wait...
echo.

REM Try curl first (Windows 10+)
curl -L -o "%INSTALLER%" "%PYTHON_URL%" 2>nul
if %errorlevel% neq 0 (
    REM Fallback to certutil
    certutil -urlcache -split -f "%PYTHON_URL%" "%INSTALLER%" >nul 2>&1
)

if not exist "%INSTALLER%" (
    echo [FAIL] Download failed!
    echo.
    echo Please manually download Python:
    echo   1. Open browser: https://www.python.org/downloads/
    echo   2. Download Python 3.11+
    echo   3. IMPORTANT: Check "Add Python to PATH" when installing!
    echo   4. After install, run this script again.
    echo.
    pause
    exit /b 1
)

echo Download complete! Installing Python...
echo (This may take 1-2 minutes)
echo.

REM === Silent install Python ===
"%INSTALLER%" /quiet InstallAllUsers=0 PrependPath=1 Include_test=0 Include_doc=0

if %errorlevel% neq 0 (
    echo.
    echo [NOTE] Silent install may have failed.
    echo Starting manual installer...
    echo.
    echo !!! IMPORTANT: Check "Add Python to PATH" at the bottom !!!
    echo.
    "%INSTALLER%"
)

REM Refresh PATH
set "PATH=%LOCALAPPDATA%\Programs\Python\Python311;%LOCALAPPDATA%\Programs\Python\Python311\Scripts;%PATH%"

REM Verify installation
python --version >nul 2>&1
if %errorlevel% neq 0 (
    echo.
    echo [FAIL] Python install seems incomplete.
    echo Please restart this script after installation.
    echo If problem persists, restart your computer first.
    echo.
    pause
    exit /b 1
)

echo.
echo [OK] Python installed successfully!
python --version

:install_deps
echo.
echo Checking dependencies...
python -c "import win32com" >nul 2>&1
if %errorlevel% neq 0 (
    echo Installing pywin32...
    python -m pip install pywin32 -q
    echo [OK] pywin32 installed
) else (
    echo [OK] pywin32 ready
)

echo.
echo ========================================
echo    Starting Excel to PDF Tool...
echo ========================================
echo.

python "%~dp0main.py"

if %errorlevel% neq 0 (
    echo.
    echo [ERROR] Program error. See above for details.
    pause
)
goto :eof

:install_deps_py
echo.
echo Checking dependencies...
py -c "import win32com" >nul 2>&1
if %errorlevel% neq 0 (
    echo Installing pywin32...
    py -m pip install pywin32 -q
    echo [OK] pywin32 installed
) else (
    echo [OK] pywin32 ready
)

echo.
echo ========================================
echo    Starting Excel to PDF Tool...
echo ========================================
echo.

py "%~dp0main.py"

if %errorlevel% neq 0 (
    echo.
    echo [ERROR] Program error. See above for details.
    pause
)
