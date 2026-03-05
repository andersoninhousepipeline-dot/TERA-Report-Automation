"""
TERA Report — Font Installer
=============================
Run once after installation to copy the bundled fonts to the system font
directory so they are available application-wide.

Usage:
    python install_fonts.py
"""

import os
import sys
import shutil
import platform
import subprocess

FONTS_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), "fonts")

FONT_FILES = [
    "Calibri.ttf",
    "Calibri-Bold.ttf",
    "Calibri-Italic.ttf",
    "Calibri-BoldItalic.ttf",
    "GillSansMT-Bold.ttf",
    "SegoeUI.ttf",
    "SegoeUI-Bold.ttf",
    "DengXian.ttf",
    "DengXian_Bold.ttf",
]


def _available_fonts():
    return [f for f in FONT_FILES if os.path.exists(os.path.join(FONTS_DIR, f))]


def install_windows(fonts):
    import ctypes
    win_fonts = os.path.join(os.environ.get("WINDIR", r"C:\Windows"), "Fonts")
    installed, skipped = 0, 0
    for fname in fonts:
        src = os.path.join(FONTS_DIR, fname)
        dst = os.path.join(win_fonts, fname)
        if os.path.exists(dst):
            skipped += 1
            continue
        try:
            shutil.copy2(src, dst)
            # Notify Windows that a new font was added
            ctypes.windll.gdi32.AddFontResourceW(dst)
            installed += 1
        except PermissionError:
            print(f"  [!] Permission denied: {fname}")
            print("      Re-run this script as Administrator.")
            sys.exit(1)
    return installed, skipped


def install_linux(fonts):
    user_fonts = os.path.expanduser("~/.local/share/fonts/TERA")
    os.makedirs(user_fonts, exist_ok=True)
    installed, skipped = 0, 0
    for fname in fonts:
        src = os.path.join(FONTS_DIR, fname)
        dst = os.path.join(user_fonts, fname)
        if os.path.exists(dst):
            skipped += 1
            continue
        shutil.copy2(src, dst)
        installed += 1
    # Rebuild font cache
    try:
        subprocess.run(["fc-cache", "-f", user_fonts],
                       check=True, capture_output=True)
    except FileNotFoundError:
        pass  # fc-cache not available — fonts still copied
    return installed, skipped


def install_macos(fonts):
    user_fonts = os.path.expanduser("~/Library/Fonts")
    os.makedirs(user_fonts, exist_ok=True)
    installed, skipped = 0, 0
    for fname in fonts:
        src = os.path.join(FONTS_DIR, fname)
        dst = os.path.join(user_fonts, fname)
        if os.path.exists(dst):
            skipped += 1
            continue
        shutil.copy2(src, dst)
        installed += 1
    return installed, skipped


def main():
    fonts = _available_fonts()
    if not fonts:
        print(f"No font files found in: {FONTS_DIR}")
        sys.exit(1)

    print(f"TERA Report — Font Installer")
    print(f"Found {len(fonts)} font(s) in {FONTS_DIR}")

    system = platform.system()
    if system == "Windows":
        installed, skipped = install_windows(fonts)
    elif system == "Linux":
        installed, skipped = install_linux(fonts)
    elif system == "Darwin":
        installed, skipped = install_macos(fonts)
    else:
        print(f"Unsupported platform: {system}")
        sys.exit(1)

    print(f"\nDone — {installed} font(s) installed, {skipped} already present.")
    if installed:
        print("You may need to restart the application for changes to take effect.")


if __name__ == "__main__":
    main()
