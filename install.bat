@echo off
echo ============================================
echo  TERA Report Generator - First-Time Setup
echo ============================================
echo.

:: Check Python
python --version >nul 2>&1
if %errorlevel% neq 0 (
    echo [ERROR] Python not found.
    echo.
    echo Please install Python 3.10 or higher from:
    echo   https://www.python.org/downloads/
    echo.
    echo IMPORTANT: During installation, check the box that says
    echo   "Add Python to PATH"
    echo.
    pause
    exit /b 1
)

:: Check Python version is 3.10+
python -c "import sys; exit(0 if sys.version_info >= (3,10) else 1)" >nul 2>&1
if %errorlevel% neq 0 (
    echo [ERROR] Python 3.10 or higher is required.
    echo Your version:
    python --version
    echo.
    echo Download Python 3.10+ from: https://www.python.org/downloads/
    pause
    exit /b 1
)

echo [OK] Found:
python --version
echo.

:: Create virtual environment
if not exist "venv" (
    echo Creating virtual environment...
    python -m venv venv
    if %errorlevel% neq 0 (
        echo [ERROR] Failed to create virtual environment.
        pause
        exit /b 1
    )
    echo [OK] Virtual environment created.
) else (
    echo [OK] Virtual environment already exists.
)

echo.
echo Installing dependencies (this may take a few minutes)...
venv\Scripts\pip install --upgrade pip --quiet
venv\Scripts\pip install -r requirements.txt

if %errorlevel% neq 0 (
    echo.
    echo [ERROR] Failed to install one or more dependencies.
    echo Please check your internet connection and try again.
    pause
    exit /b 1
)

echo.
echo ============================================
echo  Setup complete!
echo  Double-click launch_tera.bat to start.
echo ============================================
echo.
pause
