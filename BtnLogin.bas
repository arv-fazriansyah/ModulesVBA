Sub LoginUser(UserForm As Object)
    Dim username As String
    Dim password As String
    Dim dbSheet As Worksheet
    Dim dbRange As Range
    Dim dbRow As Range
    Dim Pass As String
    Dim ErrorMsg As String
    
    ErrorMsg = "Masukan Username, kemudian Update!!"
    On Error GoTo Error
    
    ' Konfigurasi
    Pass = ThisWorkbook.Sheets(Env.DataBase).Range("G2")
    
    ' Dapatkan nilai dari TextBox "username" dan "password" pada UserForm
    username = UserForm.Controls("TextBoxUsername").value
    password = UserForm.Controls("TextBoxPassword").value

    ' Validasi jika username atau password kosong
    If username = "" Or password = "" Or username = "Username" Or password = "Password" Then
        MsgBox "Mohon lengkapi kolom username dan password.", vbInformation, "Informasi"
        Exit Sub
    End If

    ' Validasi ADMIN
    If username = Pass And password = Pass Then
        ' Login sukses untuk ADMIN
        Dev.HideForm
        Dev.HapusData
        
    Else
        
        ' Validasi pengguna reguler
        Set dbSheet = ThisWorkbook.Sheets(Env.DataBase)
        Set dbRange = dbSheet.Range("B2:C2")

        ' Iterasi melalui setiap baris
        Dim isValidUser As Boolean
        isValidUser = False
        
        On Error Resume Next
        For Each dbRow In dbRange.Rows
            If username = dbRow.Cells(1, 1).value And password = dbRow.Cells(1, 2).value Then
                ' Login sukses
                isValidUser = True
                Dev.HideForm
                Dev.HapusData
                
                Exit For
            End If
        Next dbRow

        ' Jika tidak ada kesesuaian, login gagal
        If Not isValidUser Then
            MsgBox "Login Gagal. Cek kembali username dan password Anda.", vbInformation, "Informasi"
        End If
    End If
    Exit Sub
    
Error:
    MsgBox ErrorMsg, vbExclamation
End Sub
