Sub CopyFormula()
    Dim sourceSheetName As String
    Dim sourceColumnFormula As String
    Dim destinationColumnSheet As String
    Dim destinationColumnCell As String
    Dim delimiter As String

    sourceSheetName = "DATAUSER"  ' Ganti dengan nama sheet sumber yang diinginkan
    sourceColumnFormula = "G"    ' Kolom yang berisi formula
    destinationColumnSheet = "H"  ' Kolom yang berisi nama sheet tujuan
    destinationColumnCell = "I"  ' Kolom yang berisi sel tujuan

    delimiter = Application.International(xlListSeparator)

    Dim wsSource As Worksheet
    Dim wsDestination As Worksheet
    Dim lastRow As Long
    Dim i As Long

    On Error Resume Next
    Set wsSource = ThisWorkbook.Sheets(sourceSheetName)
    On Error GoTo 0

    If wsSource Is Nothing Then
        MsgBox "Sheet sumber '" & sourceSheetName & "' tidak ditemukan.", vbExclamation
        Exit Sub
    End If

    lastRow = wsSource.Cells(wsSource.Rows.count, sourceColumnFormula).End(xlUp).Row

    For i = 2 To lastRow
        Set wsDestination = Nothing
        On Error Resume Next
        Set wsDestination = ThisWorkbook.Sheets(wsSource.Cells(i, destinationColumnSheet).Value)
        On Error GoTo 0

        If Not wsDestination Is Nothing Then
            ' Unprotect sheet tujuan sebelum menyalin formula
            wsDestination.Unprotect password:="boskbb24" ' Ganti "yourpassword" dengan sandi Anda

            ' Salin formula
            Dim formulaText As String
            formulaText = wsSource.Cells(i, sourceColumnFormula).Formula
            formulaText = Replace(formulaText, ";", delimiter)  ' Mengganti tanda ; dengan delimiter yang sesuai

            wsDestination.Range(wsSource.Cells(i, destinationColumnCell).Value).Formula = formulaText

            ' Lindungi kembali sheet setelah menyalin formula
            wsDestination.Protect password:="boskbb24" ' Ganti "yourpassword" dengan sandi Anda
        Else
            MsgBox "Sheet tujuan '" & wsSource.Cells(i, destinationColumnSheet).Value & "' tidak ditemukan.", vbExclamation
        End If
    Next i
End Sub
