Private Sub CommandButtonLogin_Click()
    LoginUser Me
End Sub

Private Sub CommandButtonRefresh_Click()
BtnUpdate.GsheetDataLoginUpdate
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
'Dev.ClosedForm
End Sub

Private Sub UserForm_Initialize()
Text.Placeholder
Text.TextVersion
End Sub
Private Sub CommandButton1_Click()
Text.TogglePasswordIcon
End Sub
Private Sub TextBoxUsername_Enter()
PH_Enter TextBoxUsername
End Sub
Private Sub TextBoxUsername_Exit(ByVal Cancel As MSForms.ReturnBoolean)
PH_Exit TextBoxUsername
End Sub
Private Sub TextBoxPassword_Enter()
PH_Enter TextBoxPassword
End Sub
Private Sub TextBoxPassword_Exit(ByVal Cancel As MSForms.ReturnBoolean)
PH_Exit TextBoxPassword
End Sub
