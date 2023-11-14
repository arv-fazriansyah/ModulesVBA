Sub HapusData()
    Dim Ws As Worksheet
    Dim SheetName As String
    Dim DataRange As String
    Dim Pass As String
    Dim WrongPasswordMsg As String
    
    On Error Resume Next
    
    ' Konfigurasi
    SheetName = Env.DataBase
    DataRange = ThisWorkbook.Sheets(SheetName).Range("H2")
    Pass = ThisWorkbook.Sheets(SheetName).Range("G2")
    
    WrongPasswordMsg = "Silahkan Update Aplikasi!"
    
    Set Ws = ThisWorkbook.Sheets(SheetName)
    
    ' Periksa apakah lembar kerja ada
    If Ws Is Nothing Then
        MsgBox "Logout Aplikasi, kemudian Update pada halaman Login!", vbExclamation
        Exit Sub
    End If
    
    ' Jika lembar kerja dilindungi, lakukan unprotect terlebih dahulu
    If Ws.ProtectContents Then
        If Pass <> "" Then
            On Error Resume Next
            Ws.Unprotect Pass
            If Ws.ProtectContents Then
                MsgBox WrongPasswordMsg, vbExclamation
                Exit Sub
            End If
        End If
    End If
    
    ' Tentukan rentang sel yang ingin dihapus
    On Error GoTo Error
    Dim DeleteRange As Range
    Set DeleteRange = Ws.Range(DataRange)
    
    ' Hapus data dari rentang sel yang ditentukan
    If Not DeleteRange Is Nothing Then
        DeleteRange.ClearContents ' Menghapus isi dari sel
    End If
    
    ' Setelah pembaruan, lindungi kembali lembar kerja jika password diberikan
    If Pass <> "" Then Ws.Protect Pass
    
    ' Simpan workbook setelah menghapus data eksternal
    ThisWorkbook.Save
    Exit Sub
    
Error:
    MsgBox WrongPasswordMsg, vbExclamation
End Sub
