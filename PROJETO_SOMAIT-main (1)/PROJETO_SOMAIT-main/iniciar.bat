@echo off
setlocal

cd /d "%~dp0"
title SOMALABS - Gestao de Devolucoes

echo ================================================
echo   SOMALABS - Setup e Inicializacao Automatica
echo ================================================
echo.

:: ── 1. Detectar Python ──────────────────────────────────────────
set "PY_CMD="

where py >nul 2>nul
if %errorlevel%==0 (
    set "PY_CMD=py -3"
    goto :python_ok
)

where python >nul 2>nul
if %errorlevel%==0 (
    set "PY_CMD=python"
    goto :python_ok
)

echo [ERRO] Python nao foi encontrado neste computador.
echo.
echo Instale o Python 3.10+ em https://www.python.org/downloads/
echo Marque a opcao "Add Python to PATH" durante a instalacao.
echo.
pause
exit /b 1

:python_ok
echo [OK] Python encontrado: %PY_CMD%

:: Verificar versao minima (3.10)
for /f "tokens=2 delims= " %%v in ('%PY_CMD% --version 2^>^&1') do set "PY_VER=%%v"
echo      Versao: %PY_VER%
echo.

:: ── 2. Criar ambiente virtual se nao existir ────────────────────
if not exist ".venv\Scripts\python.exe" (
    echo [..] Criando ambiente virtual (.venv)...
    %PY_CMD% -m venv .venv
    if errorlevel 1 (
        echo [ERRO] Falha ao criar ambiente virtual.
        pause
        exit /b 1
    )
    echo [OK] Ambiente virtual criado.
    echo.
) else (
    echo [OK] Ambiente virtual ja existe.
    echo.
)

set "PY_CMD=.venv\Scripts\python.exe"
set "PIP_CMD=.venv\Scripts\pip.exe"

:: ── 3. Instalar dependencias ────────────────────────────────────
echo [..] Verificando dependencias...
"%PIP_CMD%" install -r requirements.txt --quiet --disable-pip-version-check
if errorlevel 1 (
    echo [ERRO] Falha ao instalar dependencias.
    echo Verifique sua conexao com a internet e tente novamente.
    pause
    exit /b 1
)
echo [OK] Dependencias instaladas.
echo.

:: ── 4. Inicializar banco de dados ───────────────────────────────
echo [..] Inicializando banco de dados...
"%PY_CMD%" -c "from app.database import criar; criar()"
if errorlevel 1 (
    echo [ERRO] Falha ao inicializar o banco de dados.
    pause
    exit /b 1
)
echo [OK] Banco de dados pronto.
echo.

:: ── 5. Iniciar aplicacao ────────────────────────────────────────
echo ================================================
echo   Iniciando aplicacao...
echo   Acesse: http://127.0.0.1:5000
echo   Usuario: admin  /  Senha: azzas2026
echo   Pressione CTRL+C para encerrar.
echo ================================================
echo.

:: Abrir navegador automaticamente apos 2 segundos
start "" cmd /c "timeout /t 2 >nul && start http://127.0.0.1:5000"

"%PY_CMD%" main.py

echo.
echo Aplicacao encerrada.
pause
