#!/bin/bash
cd "$(dirname "$0")"

# Use venv if available, otherwise fall back to system Python
if [ -f "venv/bin/python3" ]; then
    venv/bin/python3 tera_report_generator.py
elif command -v python3 &>/dev/null; then
    python3 tera_report_generator.py
else
    echo "[ERROR] Python3 not found. Run install.sh first."
    read -p "Press Enter to exit..."
    exit 1
fi

if [ $? -ne 0 ]; then
    echo "The application exited with an error."
    echo "Run install.sh first if you have not done so."
    read -p "Press Enter to exit..."
fi
