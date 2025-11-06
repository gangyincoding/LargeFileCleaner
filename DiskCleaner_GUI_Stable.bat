@echo off
title Disk Space Analyzer - GUI Stable Version
color 0B

echo.
echo ===============================================================
echo                    Disk Space Analyzer
echo                   GUI Stable Version v1.0
echo ===============================================================
echo.

echo Step 1: Checking Python environment...
python --version >nul 2>&1
if errorlevel 1 (
    echo.
    echo ERROR: Python not found!
    echo.
    echo Solution:
    echo 1. Visit https://www.python.org/downloads/
    echo 2. Download and install Python 3.8 or higher
    echo 3. Check "Add Python to PATH" during installation
    echo.
    pause
    exit /b 1
)

echo [SUCCESS] Python environment check passed
echo.

echo Step 2: Checking GUI files...
if not exist disk_analyzer_gui_stable.py (
    echo [ERROR] GUI program file not found
    pause
    exit /b 1
)

echo [SUCCESS] GUI file check passed
echo.

echo Step 3: Starting GUI...
echo [INFO] This is the stable version with all fixes applied
echo.

python disk_analyzer_gui_stable.py

if errorlevel 1 (
    echo.
    echo Program encountered an error
    echo.
    echo The stable version includes:
    echo - Complete error handling
    echo - Demo mode fallback
    echo - Detailed error messages
    echo.
    pause
)

echo.
echo Program has exited
timeout /t 2 >nul