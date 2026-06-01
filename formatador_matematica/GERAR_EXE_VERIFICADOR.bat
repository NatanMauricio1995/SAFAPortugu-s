@echo off
title Gerar Abridor Verificador

echo ==========================================
echo GERAR EXE VERIFICADOR
echo ==========================================
echo.

echo Instalando PyInstaller...
py -3.12 -m pip install pyinstaller

echo.
echo Gerando EXE...
py -3.12 -m PyInstaller --noconfirm --onefile --console --name Abrir_Sistema installer.py

echo.
echo Finalizado.
echo O arquivo ficou em:
echo dist\Abrir_Sistema.exe
echo.
echo Coloque Abrir_Sistema.exe na mesma pasta do app.py.
echo.
pause
