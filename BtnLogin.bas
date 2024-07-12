Sub LoginUser(UserForm As Object)
    Dim username As String
    Dim password As String
    Dim dbSheet As Worksheet
    Dim dbRange As Range
    Dim dbRow As Range
    Dim Pass As String
    Dim ErrorMsg As String
    
    ErrorMsg = "Masukan Username, kemudian tekan Update!!"
    ErrorMsgDev = "Download ulang Aplikasi, hubungi Admin"
    
    On Error GoTo Error
    
    ' Konfigurasi
    Pass = ThisWorkbook.Sheets(Env.DataBase).Range("G2")
    
    ' Dapatkan nilai dari TextBox "username" dan "password" pada UserForm
    username = UserForm.Controls("TextBoxUsername").Value
    password = UserForm.Controls("TextBoxPassword").Value

    ' Validasi jika username kosong
    If username = "" Or username = "Username" Then
        MsgBox "Mohon isi kolom username.", vbInformation, "Informasi"
        Exit Sub
    End If
    
    ' Validasi jika password kosong
    If password = "" Or password = "Password" Then
        MsgBox "Mohon isi kolom password.", vbInformation, "Informasi"
        Exit Sub
    End If

    ' Validasi ADMIN
    If username = Pass And password = Pass Then
        ' Menambakan informasi lain
        On Error GoTo ErrorDev
        ThisWorkbook.Sheets("DEV").Range("F2").Value = username
        Dev.CekIDPerangkat
        Dev.CekIPPublik
        Dev.CekNamaKomputer
        Dev.CekVersiOffice
        Dev.CekWaktu
        ' Login sukses untuk ADMIN
        Dev.HideForm
        Dev.HapusData
        On Error Resume Next
        ' Dev.SendData
        
    Else
        
        ' Validasi pengguna reguler
        If Not Dev.TesKoneksi Then
            MsgBox "Tidak ada koneksi internet.", vbExclamation
            Exit Sub
        End If
        
        BtnUpdate.GsheetDataLogin
        Set dbSheet = ThisWorkbook.Sheets(Env.DataBase)
        Set dbRange = dbSheet.Range("A2:C2")

        ' Iterasi melalui setiap baris
        Dim isValidUser As Boolean
        isValidUser = False
        
        ' On Error Resume Next
        For Each dbRow In dbRange.Rows
            If username = dbRow.Cells(1, 1).Value And password = dbRow.Cells(1, 3).Value Then
                ' Menambakan informasi lain
                On Error GoTo ErrorDev
                ThisWorkbook.Sheets("DEV").Range("F2").Value = username
                Dev.CekIDPerangkat
                Dev.CekIPPublik
                Dev.CekNamaKomputer
                Dev.CekVersiOffice
                Dev.CekWaktu
                ' Login User sukses
                isValidUser = True
                Dev.HideForm
                Dev.HapusData
                On Error Resume Next
                ' Dev.SendData
                
                Exit For
            End If
        Next dbRow

        ' Jika tidak ada kesesuaian, login gagal
        If Not isValidUser Then
            Dev.HapusData
            MsgBox "Login Gagal. Cek kembali username dan password Anda.", vbInformation, "Informasi"
        End If
    End If
    Exit Sub
    
Error:
    MsgBox ErrorMsg, vbExclamation
    Exit Sub
ErrorDev:
    MsgBox ErrorMsgDev, vbExclamation
End Sub
