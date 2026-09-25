@echo off
setlocal
cd /d "%~dp0"

if not exist ".argus_setup_ok" call :repair

for /f "delims=" %%T in ('powershell -NoProfile -Command "[Console]::Title"') do set "ARGUS_OLDTITLE=%%T"

title ARGUS
python "%~dp0argus.py" %*

if defined ARGUS_OLDTITLE (
    title %ARGUS_OLDTITLE%
) else (
    title
)
goto :eof

:repair
echo Setup marker missing -- repairing Argus...
git remote get-url origin >nul 2>&1
if %ERRORLEVEL% EQU 0 (
    echo Syncing code with git...
    git fetch origin
    for /f "delims=" %%b in ('git rev-parse --abbrev-ref HEAD') do (
        git reset --hard origin/%%b
    )
) else (
    echo No git remote configured yet -- skipping code sync.
)

echo Installing Python dependencies...
python -m pip install -r requirements.txt -q
echo Installing Playwright's browser...
python -m playwright install chromium

echo Adding Argus to your PATH...
powershell.exe -NoProfile -Command "$dir = '%~dp0'.TrimEnd('\'); $old = [Environment]::GetEnvironmentVariable('PATH','User'); if ($old -notlike \"*$dir*\") { [Environment]::SetEnvironmentVariable('PATH', \"$old;$dir\", 'User') }"

echo done > ".argus_setup_ok"
echo Repair complete.
goto :eof
