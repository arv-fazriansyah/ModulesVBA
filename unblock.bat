@echo off
setlocal

set /p file_name=Masukkan nama file tanpa ekstensi: 

:: Periksa apakah file tersebut ada dengan ekstensi .xlsb
if exist "%file_name%.xlsb" (
  set "excel_extension=xlsb"
)

:: Periksa apakah file tersebut ada dengan ekstensi .xlsm jika belum ditemukan dengan .xlsb
if not defined excel_extension (
  if exist "%file_name%.xlsm" (
    set "excel_extension=xlsm"
  )
)

:: Jika tidak ada file yang ditemukan, keluar
if not defined excel_extension (
  echo File tidak ditemukan: "%file_name%.xlsb" atau "%file_name%.xlsm"
  exit /b
)

:: Jalankan perintah PowerShell untuk mengaktifkan Unblock pada file
powershell -command "Unblock-File -Path '%file_name%.%excel_extension%'"

:: Jalankan perintah PowerShell untuk menonaktifkan Protected View
powershell -command "Set-ItemProperty -Path 'HKCU:\Software\Microsoft\Office\*\Excel\Security' -Name 'ProtectedView' -Value 0"

:: Jalankan perintah PowerShell untuk mengaktifkan Enable Content
powershell -command "Set-ItemProperty -Path 'HKCU:\Software\Microsoft\Office\*\Excel\Security' -Name 'VBAWarnings' -Value 1"

echo Opsi 'Unblock', 'Protected View', dan 'Enable Content' telah diaktifkan untuk file: "%file_name%.%excel_extension%"
pause
