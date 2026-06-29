@echo off
echo Building Sales Report App for Windows...
echo Installing/checking dependencies...
python -m pip install -r ..\requirements.txt
python -m PyInstaller --noconfirm --windowed --onefile --name "Sales Report Generator" ..\app.py
echo Build complete. Check the 'dist' folder.
pause
