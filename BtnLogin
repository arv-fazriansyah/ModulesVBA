Sub LoginUser(UserForm As Object)
    Dim username As String
    Dim Password As String
    Dim dbSheet As Worksheet
    Dim dbRange As Range
    Dim dbRow As Range
    Dim PasswordX As String
    
    ' Konfigurasi
    PasswordX = "ADMIN"
    
    ' Dapatkan nilai dari TextBox "username" dan "password" pada UserForm
    username = UserForm.Controls("TextBoxUsername").value
    Password = UserForm.Controls("TextBoxPassword").value

    ' Validasi jika username atau password kosong
    If username = "" Or Password = "" Or username = "Username" Or Password = "Password" Then
        MsgBox "Mohon lengkapi data username dan password.", vbInformation, "Informasi"
        Exit Sub
    End If

    ' Validasi ADMIN
    If username = PasswordX And Password = PasswordX Then
        ' Login sukses untuk ADMIN
        HalamanLogin.Hide
    Else
        ' Update
        BtnUpdate.GsheetData

        ' Validasi pengguna reguler
        Set dbSheet = ThisWorkbook.Sheets("DATAUSER")
        Set dbRange = dbSheet.UsedRange

        ' Iterasi melalui setiap baris
        Dim isValidUser As Boolean
        isValidUser = False

        For Each dbRow In dbRange.Rows
            If username = dbRow.Cells(1, 1).value And Password = dbRow.Cells(1, 2).value Then
                ' Login sukses
                HalamanLogin.Hide
                isValidUser = True
                Exit For
            End If
        Next dbRow

        ' Jika tidak ada kesesuaian, login gagal
        If Not isValidUser Then
            MsgBox "Login Gagal. Cek kembali username dan password Anda.", vbInformation, "Informasi"
        End If
    End If
End Sub
