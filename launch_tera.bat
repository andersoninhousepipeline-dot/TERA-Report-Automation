@echo off
:: Use venv if available, otherwise fall back to system Python
if exist "venv\Scripts\pythonw.exe" (
    start "" venv\Scripts\pythonw.exe tera_report_generator.py
) else if exist "venv\Scripts\python.exe" (
    venv\Scripts\python.exe tera_report_generator.py
    if %errorlevel% neq 0 (
        echo.
        echo The application exited with an error.
        echo Run install.bat first if you have not done so.
        pause
    )
) else (
    python tera_report_generator.py
    if %errorlevel% neq 0 (
        echo.
        echo The application exited with an error.
        echo Run install.bat first if you have not done so.
        pause
    )
)
