
@echo off
setlocal
cd /d "%~dp0"

python --version >nul 2>&1
if not errorlevel 1 goto python_ok

py --version >nul 2>&1
if not errorlevel 1 goto py_ok

echo Python nao encontrado.
echo Instale em https://www.python.org/downloads/
pause
exit /b 1

:python_ok
set "PYTHON=python"
goto install

:py_ok
set "PYTHON=py"

:install
%PYTHON% -m pip install python-docx pyinstaller
if errorlevel 1 goto failed

%PYTHON% -m PyInstaller --noconfirm --clean --windowed --name "Caderno Faculdade" --collect-all docx caderno_faculdade.py
if errorlevel 1 goto failed

echo.
echo Executavel criado em:
echo dist\Caderno Faculdade\Caderno Faculdade.exe
pause
exit /b 0

:failed
echo.
echo A compilacao falhou. Envie as mensagens anteriores.
pause
exit /b 1

