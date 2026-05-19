@echo off
REM Allow PowerShell scripts to run for the current user
powershell -Command "Set-ExecutionPolicy -ExecutionPolicy Bypass -Scope CurrentUser -Force"

set "VENV_DIR=venv"
set "REQ_FILE=requirements.txt"

REM Flag to track whether we just created the venv
set "NEW_VENV=false"

REM Check if venv folder exists
IF NOT EXIST "%VENV_DIR%" (
    echo [INFO] Virtual environment not found. Creating one...
    python -m venv %VENV_DIR%
    IF %ERRORLEVEL% NEQ 0 (
        echo [ERROR] Failed to create virtual environment.
        pause
        exit /b 1
    )
    set "NEW_VENV=true"
)

REM Activate virtual environment
call %VENV_DIR%\Scripts\activate

REM Install requirements only if venv is new
IF "%NEW_VENV%"=="true" (
    IF EXIST "%REQ_FILE%" (
        echo [INFO] Installing dependencies from %REQ_FILE%...
        pip install --upgrade pip
        pip install -r %REQ_FILE%
    ) ELSE (
        echo [WARNING] No %REQ_FILE% found. Skipping dependency installation.
    )
) ELSE (
    echo [INFO] Existing virtual environment detected. Skipping dependency installation.
)

REM Run your Streamlit app
echo [INFO] Starting Streamlit app...
streamlit run Main.py

REM Optional: keep window open after exit
pause
