@echo off
setlocal
cd /d "%~dp0"
powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%~dp0scripts\backup-commit-push.ps1" %*
echo.
echo Selesai menjalankan script. Jika ada error di atas, baca pesannya sebelum menutup jendela ini.
pause
endlocal
