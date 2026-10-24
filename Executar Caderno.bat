@echo off
cd /d "%~dp0"

where py >nul 2>&1
if errorlevel 1 (
    echo Python nao encontrado.
    echo Instale em https://www.python.org/downloads/
    pause
    exit /b 1
)

echo Instalando dependencias...
py -m pip install --upgrade pip
if errorlevel 1 goto erro

py -m pip install -r requirements.txt
if errorlevel 1 goto erro

py -m pip install pyinstaller
if errorlevel 1 goto erro

echo.
echo Compilando o Caderno Faculdade...
py -m PyInstaller --noconfirm --clean --windowed --name "Caderno Faculdade" --collect-all docx caderno_faculdade.py
if errorlevel 1 goto erro

echo.
echo Compilacao concluida!
echo Executavel em:
echo dist\Caderno Faculdade\Caderno Faculdade.exe
echo.
pause
exit /b 0

:erro
echo.
echo ERRO durante a preparacao ou compilacao.
echo Confira as mensagens acima.
pause
exit /b 1