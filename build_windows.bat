@echo off
setlocal enabledelayedexpansion

set APP_NAME=SeismoLift
set ENTRY=SeismoLift.py
set ICON=assets\seismolift_icon.ico
set DATA_XLSX=Zonas_Sismicas_PT.xlsx

where python >nul 2>nul
if errorlevel 1 (
    echo [ERRO] Python nao encontrado no PATH.
    echo Instale Python 3.10+ ou active o ambiente virtual antes de executar este script.
    pause
    exit /b 1
)

if not exist "%ENTRY%" (
    echo [ERRO] %ENTRY% nao encontrado. Execute este script dentro da pasta SeismoLift.
    pause
    exit /b 1
)

if not exist "%DATA_XLSX%" (
    echo [ERRO] %DATA_XLSX% nao encontrado.
    pause
    exit /b 1
)

if not exist "%ICON%" (
    echo [ERRO] %ICON% nao encontrado.
    pause
    exit /b 1
)

echo [1/4] A instalar dependencias...
python -m pip install --upgrade pip
python -m pip install -r requirements-build.txt
if errorlevel 1 (
    echo [ERRO] Falha na instalacao das dependencias.
    pause
    exit /b 1
)

echo [2/4] A limpar builds anteriores...
if exist build rmdir /s /q build
if exist dist rmdir /s /q dist
if exist "%APP_NAME%.spec" del "%APP_NAME%.spec"

echo [3/4] A gerar executavel Windows...
set MODE=
if /I "%1"=="onefile" set MODE=--onefile

python -m PyInstaller --noconfirm --clean --windowed !MODE! ^
  --name "%APP_NAME%" ^
  --icon "%ICON%" ^
  --add-data "%DATA_XLSX%;." ^
  --add-data "assets;assets" ^
  --collect-all matplotlib ^
  --hidden-import openpyxl.cell._writer ^
  "%ENTRY%"

if errorlevel 1 (
    echo [ERRO] PyInstaller falhou.
    pause
    exit /b 1
)

echo [4/4] Build concluido.
if /I "%1"=="onefile" (
    echo Executavel: dist\%APP_NAME%.exe
) else (
    echo Pasta distribuivel: dist\%APP_NAME%\
    echo Executavel: dist\%APP_NAME%\%APP_NAME%.exe
)

echo.
echo Nota: para criar uma versao de ficheiro unico, execute:
echo build_windows.bat onefile
pause
