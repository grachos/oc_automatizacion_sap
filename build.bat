@echo off
cd /d %~dp0
pip install -r requirements.txt pyinstaller
python -m PyInstaller --noconfirm --onefile --windowed --name CrearOCsSAP --hidden-import win32com.client --hidden-import pythoncom app.py
if not exist dist\CrearOCsSAP.exe (echo ERROR: no se genero el exe & pause & exit /b 1)
copy /y .env.example dist\.env.example >nul
copy /y plantilla_ocs.xlsx dist\plantilla_ocs.xlsx >nul
echo.
echo Listo: dist\CrearOCsSAP.exe
pause
