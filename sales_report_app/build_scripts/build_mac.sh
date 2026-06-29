#!/bin/bash
echo "Building Sales Report App for macOS..."
python3 -m PyInstaller --noconfirm --windowed \
    --name "Sales Report Generator" \
    ../app.py

echo "Build complete. Check the 'dist' folder."
