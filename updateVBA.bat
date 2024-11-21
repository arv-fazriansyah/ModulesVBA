set "install_dir=%CD%"
set "source=%install_dir%\temp\home"
set "exe=%install_dir%\temp\zip\7-Zip.exe"
set "backup_dir=%install_dir%\backup"
set "file="

:: Mencari file Excel (.xlsb) di direktori instalasi
for %%i in ("%install_dir%\*.xlsb") do (
    set "file=%install_dir%\%%~nxi"
    goto :file_found
)

echo Simpan terlebih dahulu file RBK disini %install_dir%
echo.
goto :end

:file_found

:: Membuat folder backup jika belum ada
if not exist "%backup_dir%" (
    mkdir "%backup_dir%"
)

:: Membackup file Excel ke folder backup
echo Membackup file: %file% ke folder backup...
xcopy "%file%" "%backup_dir%\" /Y

:: Mengecek apakah 7-Zip terpasang
IF EXIST "%ProgramFiles%\7-Zip\7z.exe" (
    rem echo 7-Zip sudah terpasang.
) ELSE (
    echo 7-Zip belum terpasang. Sedang menginstal...
    echo.
    :: Instalasi 7-Zip dalam mode diam
    "%exe%" /S
    echo 7-Zip telah terinstal.
    echo.
)

:: Proses kompresi file menggunakan 7-Zip
"%ProgramFiles%\7-Zip\7z.exe" a "%file%" "%source%\*"
echo File Berhasil diupdate!
echo.

:end
pause
