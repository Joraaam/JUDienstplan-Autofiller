@echo off
rem Baut Dienstplan-Autofill.exe fuer Windows (ohne Python beim Empfaenger lauffaehig).
cd /d "%~dp0"
python -m pip install -r requirements.txt pyinstaller
python -m PyInstaller --noconfirm --onefile --name Dienstplan-Autofill --add-data "public;public" --add-data "core/k4_formula.txt;core" --collect-all pdfminer api/index.py
echo.
echo Fertig: dist\Dienstplan-Autofill.exe
pause
