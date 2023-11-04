@echo off
setlocal

set /p file_name=Masukkan nama file tanpa ekstensi: 

:: Periksa apakah file tersebut ada
if not exist "%file_name%.xlsb" (
  echo File tidak ditemukan: "%file_name%.xlsb"
  exit /b
)

:: Jalankan perintah PowerShell untuk mengaktifkan Unblock pada file
powershell -command "Unblock-File -Path '%file_name%.xlsb'"

:: Jalankan perintah PowerShell untuk menonaktifkan Protected View
powershell -command "Set-ItemProperty -Path 'HKCU:\Software\Microsoft\Office\16.0\Excel\Security' -Name 'ProtectedView' -Value 0"

:: Jalankan perintah PowerShell untuk mengaktifkan Enable Content
powershell -command "Set-ItemProperty -Path 'HKCU:\Software\Microsoft\Office\16.0\Excel\Security' -Name 'VBAWarnings' -Value 1"

echo Opsi 'Unblock', 'Protected View', dan 'Enable Content' telah diaktifkan untuk file: "%file_name%.xlsb"
pause
