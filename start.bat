@echo off
cd /d "%~dp0"

echo Installing Python dependencies...
python -m pip install -r requirements.txt -q

echo Installing Playwright's browser (needed by 3 of the tools)...
python -m playwright install chromium

echo Adding Argus to your PATH so 'argus' works from any folder...
powershell.exe -NoProfile -Command "$dir = '%~dp0'.TrimEnd('\'); $old = [Environment]::GetEnvironmentVariable('PATH','User'); if ($old -notlike \"*$dir*\") { [Environment]::SetEnvironmentVariable('PATH', \"$old;$dir\", 'User'); Write-Output 'PATH updated.' } else { Write-Output 'Already in PATH.' }"

echo done > .argus_setup_ok

echo.
echo Setup complete.
echo Close this terminal and open a NEW one, then just type: argus
