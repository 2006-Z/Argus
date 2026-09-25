@echo off
cd /d "%~dp0"
echo Checking dependencies (first run only, may take a while)...
python -m pip install -r requirements.txt -q
python -m playwright install chromium
python argus.py
pause
