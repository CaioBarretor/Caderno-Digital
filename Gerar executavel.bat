@echo off
setlocal
cd /d "%~dp0"
title Gerar Caderno Faculdade

where py >nul 2>&1
if not errorlevel 1 (
    set "PYTHON=py"
    goto python_ok
)
where python >nul 2>&1
if not errorlevel 1 (
    set "PYTHON=python"
    goto python_ok
)

echo Python nao foi encontrado.
echo Instale o Python 3.12 ou superior em https://www.python.org/downloads/
echo Durante a instalacao, marque "Add Python to PATH".
pause
exit /b 1

:python_ok
%PYTHON% -m pip install --upgrade pip
if errorlevel 1 goto failed
%PYTHON% -m pip install -r requirements.txt pyinstaller
if errorlevel 1 goto failed
%PYTHON% -m PyInstaller --noconfirm --clean --onefile --windowed --name "Caderno Faculdade" --collect-all docx launcher.py
if errorlevel 1 goto failed

echo.
echo Executavel criado em:
echo dist\Caderno Faculdade.exe
pause
exit /b 0

:failed
echo.
echo A geracao falhou. Copie e envie as mensagens de erro acima.
pause
exit /b 1
