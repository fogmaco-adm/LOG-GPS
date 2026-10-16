@echo off
title Upload Timesheet ke Google Drive

:: Masuk otomatis ke folder tempat file .bat ini berada (00. PORTAL TIMESHEET)
cd /d "%~dp0"

echo ========================================================
echo       MENJALANKAN SKRIP UPLOAD TIMESHEET OTOMATIS
echo ========================================================
echo.

:: Menjalankan python portable (tanpa simbol & di depan)
"d:\FILE FOG\03. 2026\18. FOG PROJECT\Python 314\python.exe" "upload_to_drive.py"

echo.
echo ========================================================
echo Proses selesai! Tekan tombol apa saja untuk menutup.
echo ========================================================
pause