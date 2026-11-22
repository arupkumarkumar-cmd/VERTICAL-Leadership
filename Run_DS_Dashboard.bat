@echo off
title DS Weekly Operations Dashboard

cd /d "%~dp0"

echo ===================================================
echo           DS WEEKLY OPERATIONS DASHBOARD           
echo ===================================================
echo.

set "PYTHON_EXE="

if exist "%~dp0venv\Scripts\python.exe" (
    set "PYTHON_EXE=%~dp0venv\Scripts\python.exe"
    goto :PYTHON_FOUND
)

if exist "%LOCALAPPDATA%\anaconda3\python.exe" (
    set "PYTHON_EXE=%LOCALAPPDATA%\anaconda3\python.exe"
    goto :PYTHON_FOUND
)

if exist "%USERPROFILE%\anaconda3\python.exe" (
    set "PYTHON_EXE=%USERPROFILE%\anaconda3\python.exe"
    goto :PYTHON_FOUND
)

where python >nul 2>&1
if %ERRORLEVEL% EQU 0 (
    set "PYTHON_EXE=python"
    goto :PYTHON_FOUND
)

echo [ERROR] No Python interpreter found!
echo Please install Python 3.10+ or Anaconda, or configure a virtual environment.
goto :END

:PYTHON_FOUND
echo [*] Using Python interpreter: "%PYTHON_EXE%"

"%PYTHON_EXE%" -c "import pandas, openpyxl" >nul 2>&1
if %ERRORLEVEL% NEQ 0 (
    echo [*] Required packages not found. Attempting to install from requirements.txt...
    "%PYTHON_EXE%" -m pip install -r requirements.txt
    if %ERRORLEVEL% NEQ 0 (
        echo [ERROR] Failed to verify or install required packages.
        echo Please ensure pandas and openpyxl are installed.
        goto :END
    )
)

echo [*] Starting dashboard application...
echo.
"%PYTHON_EXE%" -m src.main

:END
echo.
echo ===================================================
echo Processing completed or terminated.
echo Check the logs folder for detailed execution logs.
echo ===================================================
echo.
pause
