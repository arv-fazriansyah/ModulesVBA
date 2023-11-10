Dim PasswordVisible As Boolean

Public Sub Placeholder()
    ' Set teks petunjuk dan warna untuk TextBox "Username"
    With HalamanLogin.TextBoxUsername
        .Tag = "Username"
        .value = .Tag
        .ForeColor = RGB(169, 169, 169) ' Atur warna teks abu-abu
    End With

    ' Set teks petunjuk dan warna untuk TextBox "Password"
    With HalamanLogin.TextBoxPassword
        .Tag = "Password"
        .value = .Tag
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
    If ctrl.value = ctrl.Tag Then
        ctrl.value = ""
    End If
End Sub

Public Sub PH_Exit(ctrl As MSForms.TextBox)
    If Len(ctrl.value) = 0 Then
        ctrl.value = ctrl.Tag
    End If
End Sub

Public Sub TogglePasswordIcon_Click()
    PasswordVisible = Not PasswordVisible
    TogglePasswordVisibility
End Sub
