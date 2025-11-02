@echo off
REM filepath: c:\Users\dgonzalez\Desktop\docxTopdf\build_pyinstaller.cmd
REM Genera un executable amb PyInstaller

echo Instal·lant PyInstaller...
python -m pip install --upgrade pyinstaller

echo.
echo Generant executable...
pyinstaller --noconfirm --clean --onefile --windowed ^
  --name docx-to-pdf-zip ^
  --paths src ^
  --hidden-import win32com.client ^
  --hidden-import psutil ^
  src\docx_to_pdf_zip_app.py

IF ERRORLEVEL 1 (
  echo [ERROR] PyInstaller ha fallat.
  pause
  exit /b 1
)

echo.
echo [SUCCESS] Executable generat a: dist\docx-to-pdf-zip.exe
echo Prova executar-lo i comprova la consola per errors.
pause