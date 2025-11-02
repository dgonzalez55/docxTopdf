# Convertidor DOCX a PDF amb ZIP Protegit

Aplicació amb interfície gràfica per convertir múltiples fitxers DOCX a PDF i empaquetar-los en un fitxer ZIP protegit amb contrasenya. Suporta conversió paral·lela (fins a 16 fils), reintents agressius i informes detallats. Dissenyada per a Windows.

## ✨ Característiques

- Conversió paral·lela configurable (1-16 fils)
- Reintents automàtics amb mètode alternatiu
- Creació de ZIP amb contrasenya AES
- Informe final amb estadístiques de conversions, reintents i errors
- Interfície gràfica intuïtiva amb Tkinter
- Gestió de memòria i cancel·lació de processos

## 📋 Requisits

- Python 3.8 o superior
- Dependències: `docx2pdf`, `pyzipper`, `psutil` (instal·lades automàticament)
- Opcional: `pywin32` per mètode alternatiu de conversió

## 🚀 Instal·lació

### Instal·lació local (desenvolupament)
1. Clona o descarrega el projecte.
2. Executa `install_local.cmd` per instal·lar en mode editable.
3. Executa `docx-to-pdf-zip` des de la línia de comandes.

## 📖 Ús

### Executar l'aplicació
- Des de codi font: `run_app.cmd`
- Després d'instal·lar: `docx-to-pdf-zip`
- Amb executable: Executa `dist\docx-to-pdf-zip.exe` (generat amb PyInstaller)

### Passos a l'app
1. Selecciona fitxers DOCX.
2. Opcional: Configura contrasenya per al ZIP.
3. Ajusta el nombre de conversions paral·leles.
4. Tria destí del ZIP.
5. Inicia la conversió i espera l'informe final.

## 🛠️ Construcció d'executable

### Executable independent (Windows)
Executa `build_pyinstaller.cmd` per crear `dist\docx-to-pdf-zip.exe`.
Aquest executable inclou totes les dependències i amaga la consola.

## 📄 Llicència
MIT License. Lliure per a ús educatiu i personal. Contribucions benvingudes!

## 🧑‍💻 Autor
David González - [GitHub](https://github.com/dgonzalez55)