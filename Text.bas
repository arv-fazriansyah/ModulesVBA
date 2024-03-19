Dim PasswordVisible As Boolean

Public Sub Placeholder()
    ' Set teks petunjuk dan warna untuk TextBox "Username"
    With HalamanLogin.TextBoxUsername
        .Tag = "Username"
        .Value = .Tag
        .ForeColor = RGB(169, 169, 169) ' Atur warna teks abu-abu
    End With

    ' Set teks petunjuk dan warna untuk TextBox "Password"
    With HalamanLogin.TextBoxPassword
        .Tag = "Password"
        .Value = .Tag
        .ForeColor = RGB(169, 169, 169) ' Atur warna teks abu-abu
        PasswordVisible = True ' Defaultnya mata terbuka
        TogglePasswordVisibility
    End With
End Sub

Private Sub TogglePasswordVisibility()
    If PasswordVisible Then
        HalamanLogin.TextBoxPassword.PasswordChar = "*"
    Else
        HalamanLogin.TextBoxPassword.PasswordChar = ""
    End If
End Sub

Public Sub PH_Enter(ctrl As MSForms.TextBox)
    If ctrl.Value = ctrl.Tag Then
        ctrl.Value = ""
    End If
End Sub

Public Sub PH_Exit(ctrl As MSForms.TextBox)
    If Len(ctrl.Value) = 0 Then
        ctrl.Value = ctrl.Tag
    End If
End Sub

Public Sub TogglePasswordIcon()
    PasswordVisible = Not PasswordVisible
    TogglePasswordVisibility
End Sub

Public Sub TextVersion()
    On Error Resume Next
    Dim dataText As String
    dataText = ThisWorkbook.Sheets("DATAUSER").Range("E2").Value
    
    If dataText = "" Then
        dataText = "Update Aplikasi Anda!!"
    End If
    
    With HalamanLogin.LabelVersion
        .Tag = dataText
        .Caption = dataText
        .ForeColor = RGB(169, 169, 169) ' Atur warna teks abu-abu
    End With
End Sub
