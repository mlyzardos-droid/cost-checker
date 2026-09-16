@echo off
setlocal
cd /d "%~dp0"

where py >nul 2>&1
if not errorlevel 1 (
    set "PY_CMD=py"
) else (
    where python >nul 2>&1
    if not errorlevel 1 (
        set "PY_CMD=python"
    ) else (
        echo.
        echo Python is not installed or is not available in PATH.
        echo Install Python 3 from https://www.python.org/downloads/windows/
        echo During installation, enable: Add python.exe to PATH
        pause
        exit /b 1
    )
)

if not exist .venv-win (
    %PY_CMD% -m venv .venv-win
    if errorlevel 1 (
        echo Could not create the Python environment.
        pause
        exit /b 1
    )
)

call .venv-win\Scripts\activate.bat
python -m pip install --upgrade pip
if errorlevel 1 goto :error
python -m pip install -r requirements.txt
if errorlevel 1 goto :error
python -m pip install pyinstaller
if errorlevel 1 goto :error

python -m PyInstaller --noconfirm --clean --onefile --windowed --name "AI Cost Calculator" license_cost_calculator.py
if errorlevel 1 goto :error

echo.
echo Finished. The executable is in the dist folder:
echo %CD%\dist\AI Cost Calculator.exe
pause
exit /b 0

:error
echo.
echo Build failed. Read the error above and try again.
pause
exit /b 1
