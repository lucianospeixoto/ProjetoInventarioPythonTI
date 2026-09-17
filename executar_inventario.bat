@echo off
setlocal

rem ============================================================
rem  Inventario de TI - Execucao com um clique
rem  Este arquivo deve ficar na MESMA pasta do inventario_ti.py
rem ============================================================

title Inventario de TI
cd /d "%~dp0"

echo ==================================================
echo   INVENTARIO DE TI - Iniciando...
echo ==================================================
echo.

rem --- Verifica se o Python esta instalado e acessivel no PATH ---
where python >nul 2>nul
if errorlevel 1 (
    echo [ERRO] Python nao foi encontrado neste computador.
    echo.
    echo Instale o Python em https://www.python.org/downloads/
    echo Durante a instalacao, marque a opcao "Add Python to PATH".
    echo.
    pause
    exit /b 1
)

rem --- Verifica se o script existe na mesma pasta do .bat ---
if not exist "%~dp0inventario_ti.py" (
    echo [ERRO] O arquivo inventario_ti.py nao foi encontrado nesta pasta.
    echo Coloque o executar_inventario.bat na mesma pasta do inventario_ti.py
    echo.
    pause
    exit /b 1
)

echo Verificando dependencias (psutil, wmi, openpyxl)...
python -m pip install --quiet --disable-pip-version-check psutil wmi openpyxl
if errorlevel 1 (
    echo.
    echo [ERRO] Nao foi possivel instalar as dependencias.
    echo Verifique a conexao com a internet ou permissoes deste PC.
    echo.
    pause
    exit /b 1
)

echo.
echo Executando coleta de inventario...
echo.
python "%~dp0inventario_ti.py"

endlocal
