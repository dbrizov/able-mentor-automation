@echo off

setlocal
cd /d "%~dp0"

if not exist ".venv\Scripts\python.exe" (
    echo Creating the virtual environment...
    python -m venv .venv
    if errorlevel 1 goto error
)

echo Installing the packages from requirements.txt...
".venv\Scripts\python.exe" -m pip install -r requirements.txt
if errorlevel 1 goto error

echo Installing Chromium for Playwright...
".venv\Scripts\python.exe" -m playwright install chromium
if errorlevel 1 goto error

echo.
echo Done.
pause
exit /b 0

:error
echo.
echo Setup failed, see the errors above.
pause
exit /b 1
