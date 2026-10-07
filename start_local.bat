@echo off
rem Lokal starten (benoetigt Python 3.10+)
cd /d "%~dp0"
python -m pip install -r requirements.txt -q
python api/index.py
pause
