Sub Unhide()
    Application.Visible = False ' Sembunyikan jendela Excel
    HalamanLogin.Show ' Tampilkan UserForm login
End Sub

Sub Hide()
    Application.Visible = True ' Tampilkan jendela Excel
    HalamanLogin.Hide ' Sembunyikan UserForm login
End Sub
