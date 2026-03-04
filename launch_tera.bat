@echo off
echo Starting TERA Report Generator...

:: Check if Python is installed
python --version >nul 2>&1
if %errorlevel% neq 0 (
    echo Python is not installed or not in your PATH. Please install Python.
    pause
    exit /b 1
)

:: Run the application
python tera_report_generator.py

if %errorlevel% neq 0 (
    echo The application exited with an error.
    pause
)
