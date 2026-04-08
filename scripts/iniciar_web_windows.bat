@echo off
setlocal

cd /d "%~dp0\.."
title SOMALABS - Gestao de Devolucoes

echo ================================================
echo   SOMALABS - Inicializador Web
echo ================================================
echo.

set "PY_CMD="

if exist ".venv\Scripts\python.exe" (
    set "PY_CMD=.venv\Scripts\python.exe"
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
