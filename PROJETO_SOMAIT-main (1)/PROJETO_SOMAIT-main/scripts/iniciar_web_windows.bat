@echo off
setlocal

cd /d "%~dp0\.."
title SOMALABS - Gestao de Devolucoes

echo ================================================
echo   SOMALABS - Inicializador Web
echo ================================================
echo.

set "PY_CMD="
set "VENV_LOCAL=.venv\Scripts"
set "VENV_PARENT=..\.venv\Scripts"

)if exist "%VENV_LOCAL%\python.exe" (
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

echo Usando interpretador: %PY_CMD%
echo.

echo Iniciando servidor em http://127.0.0.1:5000 ...
start "" cmd /c "timeout /t 2 >nul && start http://127.0.0.1:5000"

%PY_CMD% -m app.web

echo.
echo Servidor encerrado.
pause
