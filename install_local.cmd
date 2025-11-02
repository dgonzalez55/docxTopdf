@echo off
REM filepath: c:\Users\dgonzalez\Desktop\docxTopdf\install_local.cmd
REM Instal·la el projecte localment en mode editable

echo Instal·lant dependències...
python -m pip install --upgrade pip
python -m pip install -e .

IF ERRORLEVEL 1 (
  echo [ERROR] Instal·lació fallida.
  pause
  exit /b 1
)

echo.
echo [SUCCESS] Paquet instal·lat en mode editable.
echo Pots executar: docx-to-pdf-zip
pause