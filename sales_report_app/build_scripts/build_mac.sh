#!/bin/bash
echo "Building Sales Report App for macOS..."
pyinstaller --noconfirm --windowed \
    --name "Sales Report Generator" \
    ../app.py

echo "Build complete. Check the 'dist' folder."
