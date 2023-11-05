@echo off
setlocal

:: Hentikan semua instansi Excel terlebih dahulu
taskkill /f /im excel.exe

:: Hapus folder pengaturan Excel untuk setiap versi
rmdir /s /q "%USERPROFILE%\AppData\Local\Microsoft\Excel"
rmdir /s /q "%USERPROFILE%\AppData\Roaming\Microsoft\Excel"

:: Hapus Registry Keys Excel untuk setiap versi
reg delete "HKEY_CURRENT_USER\Software\Microsoft\Office\*.*\Excel" /f
reg delete "HKEY_CURRENT_USER\Software\Microsoft\Office\Excel" /f

:: Restart layanan Excel
start excel.exe

echo Pengaturan Excel telah direset ke default untuk semua versi Excel yang ditemukan.
pause
