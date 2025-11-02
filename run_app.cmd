@echo off
REM filepath: c:\Users\dgonzalez\Desktop\docxTopdf\run_app.cmd
REM Executa l'aplicació directament des del codi font

cd /d "%~dp0"
python src\docx_to_pdf_zip_app.py
pause