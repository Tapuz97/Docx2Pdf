@echo off
echo [🔧] Activating virtual environment...
call venv\Scripts\activate.bat

echo [🚀] Building standalone .exe with PyInstaller...
cd src
pyinstaller Docx2pdf_GUI.py ^
  --onefile ^
  --windowed ^
  --clean ^
  --noconfirm ^
  --icon=Docx2PDF_logo.ico ^
  --add-binary "C:/Python312/python312.dll;." ^
  --add-data "C:/Python312/Lib/site-packages/tkinterdnd2/tkdnd;tkinterdnd2/tkdnd"

echo [📦] Build complete. Copying to Desktop...
copy dist\Docx2pdf_GUI.exe %USERPROFILE%\Desktop\Docx2pdf_GUI.exe >nul

echo [✅] All done! You can now run it from your desktop.
pause
