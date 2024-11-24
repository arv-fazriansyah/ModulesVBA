Option Explicit

Dim objShell, intPingResult, strMessage
Set objShell = CreateObject("WScript.Shell")

' Coba ping ke google.com untuk cek koneksi internet
intPingResult = objShell.Run("cmd /c ping -n 1 google.com", 0, True)

' Tentukan pesan berdasarkan hasil ping
If intPingResult = 0 Then
    strMessage = "Koneksi internet tersedia."
Else
    strMessage = "Tidak ada koneksi internet."
End If

' Tampilkan pesan dengan MsgBox sistem modal
MsgBox strMessage, vbSystemModal + vbInformation, "Cek Koneksi Internet"
