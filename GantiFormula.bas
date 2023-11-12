Sub CopyFormula()
    Dim sourceSheetName As String
    Dim sourceColumnFormula As String
    Dim destinationColumnSheet As String
    Dim destinationColumnCell As String

    sourceSheetName = "DATAUSER"  ' Ganti dengan nama sheet sumber yang diinginkan
    sourceColumnFormula = "H"    ' Kolom yang berisi formula
    destinationColumnSheet = "I"  ' Kolom yang berisi nama sheet tujuan
    destinationColumnCell = "J"  ' Kolom yang berisi sel tujuan

    Dim wsSource As Worksheet
    Dim wsDestination As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim wrongPasswordMsg As String
    Dim PasswordX As String
    
    wrongPasswordMsg = "Kata sandi yang Anda berikan salah."
    PasswordX = ""

    On Error Resume Next
    Set wsSource = ThisWorkbook.Sheets(sourceSheetName)
    On Error GoTo 0

    If wsSource Is Nothing Then
        MsgBox "Lembar kerja sumber '" & sourceSheetName & "' tidak ditemukan.", vbExclamation
        Exit Sub
    End If

    lastRow = wsSource.Cells(wsSource.Rows.Count, sourceColumnFormula).End(xlUp).Row

    ' Mendapatkan pemisah yang digunakan dalam Excel
    Dim listSeparator As String
    listSeparator = Application.International(xlListSeparator)

    For i = 2 To lastRow
        Set wsDestination = Nothing
        On Error Resume Next
        Set wsDestination = ThisWorkbook.Sheets(wsSource.Cells(i, destinationColumnSheet).Value)
        On Error GoTo 0

        If Not wsDestination Is Nothing Then
            ' Jika lembar kerja dilindungi, lakukan unprotect terlebih dahulu
            If wsDestination.ProtectContents Then
                If PasswordX <> "" Then
                    On Error Resume Next
                    wsDestination.Unprotect PasswordX
                    On Error GoTo 0
                    If wsDestination.ProtectContents Then
                        MsgBox wrongPasswordMsg, vbExclamation
                        Exit Sub
                    End If
                Else
                    MsgBox "Kata sandi diperlukan untuk membuka proteksi lembar kerja.", vbExclamation
                    Exit Sub
                End If
            End If

            ' Salin formula dengan memeriksa pemisah yang digunakan
            Dim formula As String
            If listSeparator = ";" Then
                ' Jika pemisah adalah titik koma
                formula = Replace(wsSource.Cells(i, sourceColumnFormula).Formula, ",", ";")
            Else
                ' Jika pemisah adalah koma
                formula = Replace(wsSource.Cells(i, sourceColumnFormula).Formula, ";", ",")
            End If

            wsDestination.Range(wsSource.Cells(i, destinationColumnCell).Value).Formula = formula

            ' Lindungi kembali sheet setelah menyalin formula
            If PasswordX <> "" Then wsDestination.Protect PasswordX
        Else
            MsgBox "Lembar kerja tujuan '" & wsSource.Cells(i, destinationColumnSheet).Value & "' tidak ditemukan.", vbExclamation
        End If
    Next i
End Sub
