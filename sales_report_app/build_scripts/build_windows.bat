@echo off
echo Building Sales Report App for Windows...
pyinstaller --noconfirm --windowed --name "Sales Report Generator" ..\app.py
echo Build complete. Check the 'dist' folder.
pause
