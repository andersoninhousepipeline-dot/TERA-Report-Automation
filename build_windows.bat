@echo off
setlocal EnableDelayedExpansion
title TERA Report - Windows Build

echo.
echo ============================================================
echo   TERA Report Generator - Windows Installer Builder
echo ============================================================
echo.

:: ── 1. Check Python ───────────────────────────────────────────
python --version >nul 2>&1
if errorlevel 1 (
    echo [ERROR] Python is not installed or not in PATH.
    echo         Download from https://www.python.org/downloads/
    pause & exit /b 1
)
for /f "tokens=*" %%v in ('python --version') do echo [OK] %%v

:: ── 2. Install / upgrade dependencies ────────────────────────
echo.
echo [1/3] Installing Python dependencies...
pip install --upgrade pip --quiet
pip install -r requirements.txt --quiet
if errorlevel 1 (
    echo [ERROR] pip install failed. Check your internet connection.
    pause & exit /b 1
)
pip install pyinstaller --quiet
echo [OK] Dependencies installed.

:: ── 3. Build with PyInstaller ─────────────────────────────────
echo.
echo [2/3] Building application with PyInstaller...
pyinstaller TERA_Report.spec --clean --noconfirm
if errorlevel 1 (
    echo [ERROR] PyInstaller build failed.
    pause & exit /b 1
)
echo [OK] Build complete.  Output: dist\TERA Report\

:: ── 4. Optional: build Inno Setup installer ──────────────────
echo.
echo [3/3] Looking for Inno Setup to create installer...
set ISCC=
for %%p in (
    "C:\Program Files (x86)\Inno Setup 6\ISCC.exe"
    "C:\Program Files\Inno Setup 6\ISCC.exe"
    "C:\Program Files (x86)\Inno Setup 5\ISCC.exe"
) do (
    if exist %%p set ISCC=%%p
)

if defined ISCC (
    echo [OK] Inno Setup found: %ISCC%
    %ISCC% installer.iss
    if errorlevel 1 (
        echo [WARN] Installer creation failed, but the app is ready in dist\TERA Report\
    ) else (
        echo [OK] Installer created: Output\TERA_Report_Setup.exe
    )
) else (
    echo [SKIP] Inno Setup not found.
    echo        The app folder is ready at: dist\TERA Report\
    echo        To create a single-click .exe installer, install Inno Setup 6 from:
    echo        https://jrsoftware.org/isdl.php
    echo        then re-run this script.
)

echo.
echo ============================================================
echo   Done!
echo   App folder : dist\TERA Report\TERA Report.exe
if defined ISCC echo   Installer  : Output\TERA_Report_Setup.exe
echo ============================================================
echo.
pause
