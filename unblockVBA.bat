@echo off
setlocal enabledelayedexpansion

:: Mencari file dengan ekstensi .xlsb atau .xlsm
for %%f in (*.xlsb *.xlsm) do (
  set "excel_extension=%%~xf"
  set "file_name=%%~nf"
  powershell -command "Unblock-File -Path '!file_name!!excel_extension!'"
  powershell -command "Set-ItemProperty -Path 'HKCU:\Software\Microsoft\Office\*\Excel\Security' -Name 'ProtectedView' -Value 0"
  powershell -command "Set-ItemProperty -Path 'HKCU:\Software\Microsoft\Office\*\Excel\Security' -Name 'VBAWarnings' -Value 1"
  echo File: "!file_name!!excel_extension!" diaktifkan.
)

:: Jika tidak ada file yang ditemukan, keluar
if not defined excel_extension (
  echo Tidak ada file .xlsb atau .xlsm ditemukan.
  exit /b
)

pause
