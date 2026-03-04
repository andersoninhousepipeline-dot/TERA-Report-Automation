#!/bin/bash
echo "Starting TERA Report Generator..."

# Check if Python is installed
if ! command -v python3 &> /dev/null; then
    echo "Python3 is not installed or not in your PATH. Please install Python3."
    exit 1
fi

# Run the application
python3 tera_report_generator.py

if [ $? -ne 0 ]; then
    echo "The application exited with an error."
    read -p "Press Enter to continue..."
fi
