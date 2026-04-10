@echo off
setlocal EnableDelayedExpansion

:: Garante que o diretorio de trabalho seja o da .bat (mesmo com espacos/parenteses no caminho)
cd /d "%~dp0"
title SOMALABS - Gestao de Devolucoes

echo ================================================
echo   SOMALABS - Setup e Inicializacao Automatica
echo ================================================
echo.

:: ── 1. Detectar Python ──────────────────────────────────────────
set "PY_CMD="

where py >nul 2>nul
if !errorlevel!==0 (
    set "PY_CMD=py -3"
    goto :python_ok
)

where python >nul 2>nul
if !errorlevel!==0 (
    set "PY_CMD=python"
    goto :python_ok
)

echo [ERRO] Python nao foi encontrado neste computador.
echo.
echo Instale o Python 3.10+ em https://www.python.org/downloads/
echo Marque a opcao "Add Python to PATH" durante a instalacao.
echo.
goto :falha

:python_ok
echo [OK] Python encontrado: !PY_CMD!

:: Verificar versao
for /f "tokens=2 delims= " %%v in ('!PY_CMD! --version 2^>^&1') do set "PY_VER=%%v"
echo      Versao: !PY_VER!
echo.

:: ── 2. Criar ambiente virtual se nao existir ────────────────────
if not exist ".venv\Scripts\python.exe" (
    echo [..] Criando ambiente virtual ^(.venv^)...
    !PY_CMD! -m venv .venv
    if !errorlevel! neq 0 (
        echo [ERRO] Falha ao criar ambiente virtual.
        goto :falha
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
"!PIP_CMD!" install -r requirements.txt --quiet --disable-pip-version-check 2>&1
if !errorlevel! neq 0 (
    echo.
    echo [ERRO] Falha ao instalar dependencias.
    echo Verifique sua conexao com a internet e tente novamente.
    goto :falha
)
echo [OK] Dependencias instaladas.
echo.

:: ── 4. Inicializar banco de dados ───────────────────────────────
echo [..] Inicializando banco de dados...
"!PY_CMD!" -c "from app.database import criar; criar()" 2>&1
if !errorlevel! neq 0 (
    echo [ERRO] Falha ao inicializar o banco de dados.
    goto :falha
)
echo [OK] Banco de dados pronto.
echo.

:: ── 5. Iniciar aplicacao (modo web) ─────────────────────────────
set "FLASK_HOST=127.0.0.1"
set "FLASK_PORT=5000"

echo ================================================
echo   Iniciando aplicacao...
echo   Acesse: http://127.0.0.1:5000
echo   Usuario: admin  /  Senha: azzas2026
echo   Pressione CTRL+C para encerrar.
echo ================================================
echo.

:: Abrir navegador automaticamente apos 2 segundos
start "" cmd /c "timeout /t 2 >nul && start http://127.0.0.1:5000"

:: Iniciar direto pelo Flask (evita problemas com PySide6/Qt)
"!PY_CMD!" -m app.web

echo.
echo Aplicacao encerrada.
pause
exit /b 0

:falha
echo.
echo ================================================
echo   Pressione qualquer tecla para fechar...
echo ================================================
pause >nul
exit /b 1
pause
