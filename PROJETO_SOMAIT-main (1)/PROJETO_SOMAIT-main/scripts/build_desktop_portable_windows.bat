@echo off
setlocal

cd /d "%~dp0\.."
title SOMALABS - Build Desktop Portatil

echo ================================================
echo   SOMALABS - Build Desktop Portatil
echo ================================================
echo.

set "PY_CMD="
set "VENV_LOCAL=.venv\Scripts"
set "VENV_PARENT=..\.venv\Scripts"
set "PORTABLE_DIR=dist\SOMALABSDesktop-portable"

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

echo Usando interpretador: %PY_CMD%
echo.
echo Gerando executavel com PyInstaller...
echo.

%PY_CMD% -m PyInstaller --noconfirm --clean scripts\main.spec
if errorlevel 1 (
    echo.
    echo Falha ao gerar o executavel.
    exit /b 1
)

if exist "%PORTABLE_DIR%" (
    rmdir /s /q "%PORTABLE_DIR%"
)

mkdir "%PORTABLE_DIR%"
copy /y "dist\SOMALABSDesktop.exe" "%PORTABLE_DIR%\SOMALABSDesktop.exe" >nul

if exist "config.json" copy /y "config.json" "%PORTABLE_DIR%\config.json" >nul
if exist "database.db" copy /y "database.db" "%PORTABLE_DIR%\database.db" >nul
if exist "assets" xcopy "assets" "%PORTABLE_DIR%\assets\" /E /I /Y >nul
if exist "uploads" xcopy "uploads" "%PORTABLE_DIR%\uploads\" /E /I /Y >nul

echo.
echo Build concluido.
echo Pasta portatil: %PORTABLE_DIR%
echo.
pause