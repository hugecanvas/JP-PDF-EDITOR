@echo off
cd /d "%~dp0"
echo [1/2] Installing build requirements...
pip install "pymupdf==1.25.5" fonttools pyinstaller
echo [2/2] Building single-file exe...
pyinstaller --noconsole --onefile --clean --name "JP-PDF EDITOR" --icon jp_pdf.ico --version-file version_info.txt --add-data "jp_pdf.ico;." --hidden-import fontTools.subset --hidden-import fontTools.ttLib --exclude-module PIL --exclude-module numpy pdf_editor.py
echo.
echo Done. See dist\JP-PDF EDITOR.exe
pause
