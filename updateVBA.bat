@echo off
set "install_dir=%CD%"
set "source=%install_dir%\temp\home"
set "exe=%install_dir%\temp\zip\7-Zip.exe"
set "backup_dir=%install_dir%\backup"
set "file="
set "original_name="

:: Mengecek koneksi internet dengan ping
echo Mengecek koneksi internet...
ping -n 1 google.com >nul 2>nul
if errorlevel 1 (
    set message=Tidak ada koneksi internet. Silakan periksa koneksi Anda.
    call :msg
    goto :cleanup
)

:: Lanjutkan proses download jika ada koneksi
curl -L https://github.com/arv-fazriansyah/updateVBA/archive/refs/heads/main.zip -o updateVBA.zip || goto :error
tar -xf updateVBA.zip --strip-components=1 "updateVBA-main/*" || goto :error
del updateVBA.zip

:: Menyembunyikan folder temp setelah ekstraksi
attrib +h "%install_dir%\temp"

:: Mencari file Excel (.xlsb) di direktori instalasi
for %%i in ("%install_dir%\*.xlsb") do (
    set "file=%install_dir%\%%~nxi"
    set "original_name=%%~nxi"
    goto :file_found
)

:: Notify if no Excel file is found
set message=Simpan terlebih dahulu file RBK disini: %install_dir%
call :msg
goto :cleanup

:file_found

:: Membuat folder backup jika belum ada
if not exist "%backup_dir%" (
    mkdir "%backup_dir%"
)

:: Membackup file Excel ke folder backup
echo Membackup file: 
xcopy "%file%" "%backup_dir%\" /Y

:: Notify that backup is complete
set message=File berhasil dibackup ke folder backup.
REM call :msg

:: Mengecek apakah 7-Zip terpasang
IF EXIST "%ProgramFiles%\7-Zip\7z.exe" (
    rem echo 7-Zip sudah terpasang.
) ELSE (
    echo 7-Zip belum terpasang. Sedang menginstal...
    echo.
    :: Instalasi 7-Zip dalam mode diam
    "%exe%" /S || goto :error
    :: Notify that 7-Zip has been installed
    set message=7-Zip telah terinstal.
    REM call :msg
)

:: Proses kompresi file menggunakan 7-Zip
start /min "" "%ProgramFiles%\7-Zip\7z.exe" a "%file%" "%source%\*" || goto :error

:: Notify that the file update was successful
set message=File berhasil diupdate!
call :msg

:: Rename file setelah update
set "new_name=update_%original_name%"
ren "%file%" "%new_name%" || goto :error

:cleanup
:: Menghapus folder temp setelah selesai atau jika ada error
if exist "%install_dir%\temp" (
    rmdir /s /q "%install_dir%\temp"
)

:end
exit

:error
set message=Terjadi kesalahan.
call :msg
goto :cleanup

:msg
:: Create and run a VBS script for the message box and sound
set tempPath=%temp%\msgbox.vbs
echo Set objShell = CreateObject("WScript.Shell") > %tempPath%
echo objShell.Popup "%message%", 0, "Pemberitahuan", 64 + 4096 >> %tempPath%
echo objShell.SoundPlay "SystemHand" >> %tempPath%
cscript //nologo %tempPath%
del %tempPath%
goto :eof
