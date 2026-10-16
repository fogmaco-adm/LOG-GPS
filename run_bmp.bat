@echo off
echo Sedang memproses Timesheet BMP...
cd /d "%~dp0"
"d:\FILE FOG\03. 2026\18. FOG PROJECT\Python 314\python.exe" generate_pdf_bmp.py
echo Selesai!
pause