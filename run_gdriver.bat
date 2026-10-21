@echo off
title Upload Timesheet ke Google Drive

:: Masuk otomatis ke folder tempat file .bat ini berada (00. PORTAL TIMESHEET)
cd /d "%~dp0"

:: Menjalankan python portable (tanpa simbol & di depan)
"d:\FILE FOG\03. 2026\18. FOG PROJECT\Python 314\python.exe" "upload_to_drive.py"

echo.
echo ====================================================
echo   MENGUNGGAH PDF BARU KE GITHUB & PORTAL ADMIN...
echo ====================================================

:: --- AUTO PUSH KE GITHUB SETELAH GENERATE PDF ---
git add 3_Output_PDF/
git commit -m "Auto-update file PDF timesheet terbaru"
git push origin main

echo.
echo ====================================================
echo   Proses selesai! Tekan tombol apa saja untuk menutup.
echo ====================================================
pause