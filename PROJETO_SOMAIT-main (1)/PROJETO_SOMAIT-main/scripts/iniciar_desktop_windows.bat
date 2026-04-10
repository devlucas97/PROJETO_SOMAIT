@echo off
setlocal

cd /d "%~dp0\.."
title SOMALABS - Gestao de Devolucoes (Desktop)

echo ================================================
echo   SOMALABS - Inicializador Desktop Qt
echo ================================================
echo.

set "PY_CMD="
set "USE_CONSOLE=1"
set "VENV_LOCAL=.venv\Scripts"
set "VENV_PARENT=..\.venv\Scripts"

if /I "%~1"=="--quiet" (
    set "USE_CONSOLE=0"
)

if "%USE_CONSOLE%"=="1" (
    if exist "%VENV_LOCAL%\python.exe" (
        set "PY_CMD=%VENV_LOCAL%\python.exe"
    ) else if exist "%VENV_PARENT%\python.exe" (
        set "PY_CMD=%VENV_PARENT%\python.exe"
    ) else (
        where py >nul 2>nul
        if %errorlevel%==0 (
            set "PY_CMD=py -3"
        ) else (
            set "PY_CMD=python"
        )
    )
) else (
    if exist "%VENV_LOCAL%\pythonw.exe" (
        set "PY_CMD=%VENV_LOCAL%\pythonw.exe"
    ) else if exist "%VENV_PARENT%\pythonw.exe" (
        set "PY_CMD=%VENV_PARENT%\pythonw.exe"
    ) else if exist "%VENV_LOCAL%\python.exe" (
        set "PY_CMD=%VENV_LOCAL%\python.exe"
    ) else if exist "%VENV_PARENT%\python.exe" (
        set "PY_CMD=%VENV_PARENT%\python.exe"
    ) else (
        where py >nul 2>nul
        if %errorlevel%==0 (
            set "PY_CMD=pyw -3"
        ) else (
            set "PY_CMD=python"
        )
    )
)

echo Usando interpretador: %PY_CMD%
echo.
echo Iniciando interface desktop...
echo.

%PY_CMD% main.py

echo.
echo Aplicacao encerrada.
pause