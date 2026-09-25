@echo off
setlocal EnableExtensions
cd /d "%~dp0"

set "VENV_PY=%~dp0.venv\Scripts\python.exe"

if exist "%VENV_PY%" goto :run

echo [INFO] Creating virtual environment...
where py >nul 2>&1
if %ERRORLEVEL%==0 (
    py -3 -m venv .venv
) else (
    where python >nul 2>&1
    if %ERRORLEVEL%==0 (
        python -m venv .venv
    ) else (
        echo [ERROR] Python not found. Install Python 3.11+ and retry.
        pause
        exit /b 1
    )
)

if not exist "%VENV_PY%" (
    echo [ERROR] Failed to create .venv
    pause
    exit /b 1
)

echo [INFO] Installing dependencies...
"%VENV_PY%" -m pip install --upgrade pip
"%VENV_PY%" -m pip install -r "%~dp0requirements.txt"
if errorlevel 1 (
    echo [ERROR] pip install failed
    pause
    exit /b 1
)

:run
echo [INFO] Starting PDF to Word converter...
"%VENV_PY%" "%~dp0run.py"
set "EC=%ERRORLEVEL%"
if not "%EC%"=="0" (
    echo.
    echo [ERROR] Exit code %EC%
    pause
)
exit /b %EC%
