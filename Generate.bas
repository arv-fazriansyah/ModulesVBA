Sub ExtractData()
    Dim dataColumnA As Range
    Dim dataColumnB As Range
    Dim dataColumnC As Range
    Dim cellA As Range
    Dim cellB As Range
    Dim cellC As Range
    Dim Ws As Worksheet
    Dim outputRangeA As Range
    Dim outputRangeB As Range
    Dim outputRangeC As Range
    Dim result As String
    Dim count As Integer
    Dim outputDataA() As Variant
    Dim outputDataB() As Variant
    Dim outputDataC() As Variant
    Dim rowCount As Integer

    Set Ws = ThisWorkbook.Sheets("RAPOR1")
    Set dataColumnA = Ws.Range("B6:B100")
    Set dataColumnB = Ws.Range("G6:G100")
    Set dataColumnC = Ws.Range("I1:I100")

    Dim targetSheetName As String
    targetSheetName = "DATARAPOR1"
    On Error Resume Next
    Set outputRangeA = ThisWorkbook.Sheets(targetSheetName).Range("A1")
    Set outputRangeB = ThisWorkbook.Sheets(targetSheetName).Range("B1")
    Set outputRangeC = ThisWorkbook.Sheets(targetSheetName).Range("C1")
    On Error GoTo 0

    If outputRangeA Is Nothing Then
        ThisWorkbook.Sheets.Add(, ThisWorkbook.Sheets(ThisWorkbook.Sheets.count)).Name = targetSheetName
        Set outputRangeA = ThisWorkbook.Sheets(targetSheetName).Range("A1")
        Set outputRangeB = ThisWorkbook.Sheets(targetSheetName).Range("B1")
        Set outputRangeC = ThisWorkbook.Sheets(targetSheetName).Range("C1")
    End If

    rowCount = 1

    For Each cellA In dataColumnA.Cells
        dataA = cellA.Value
        ReDim Preserve outputDataA(1 To rowCount)
        dataB = dataColumnB.Cells(cellA.Row - 5, 1).Value ' Kolom B dimulai dari baris ke-6
        ReDim Preserve outputDataB(1 To rowCount)

        dataC = dataColumnC.Cells(cellA.Row, 1).Value
        ReDim Preserve outputDataC(1 To rowCount)

        dataArray = Split(dataC, vbLf)
        result = ""
        count = 0

        For i = LBound(dataArray) To UBound(dataArray)
            If InStr(1, dataArray(i), "- ", vbTextCompare) > 0 Then
                result = result & Trim(Mid(dataArray(i), 3)) & vbLf
                count = count + 1
                If count = 3 Then
                    Exit For ' Hanya ambil 3 kalimat pertama
                End If
            End If
        Next i

        Dim sentences() As String
        sentences = Split(result, vbLf)

        outputDataA(rowCount) = dataA
        outputDataB(rowCount) = dataB
        outputDataC(rowCount) = result ' Menyimpan 3 kalimat pertama dalam satu sel

        rowCount = rowCount + 1

        For j = LBound(sentences) To UBound(sentences)
            outputRangeA.Value = outputDataA(rowCount - 1)
            outputRangeB.Value = outputDataB(rowCount - 1)
            outputRangeC.Value = sentences(j)
            If j < UBound(sentences) Then
                Set outputRangeA = outputRangeA.Offset(1, 0)
                Set outputRangeB = outputRangeB.Offset(1, 0)
                Set outputRangeC = outputRangeC.Offset(1, 0)
            End If
        Next j

        If rowCount > 1 And cellA.Offset(1, 0).Value <> "" Then
            Set outputRangeA = outputRangeA.Offset(1, 0)
            Set outputRangeB = outputRangeB.Offset(1, 0)
            Set outputRangeC = outputRangeC.Offset(1, 0)
            rowCount = rowCount + 1
        End If
    Next cellA

    ThisWorkbook.Sheets(targetSheetName).Range("A1").Resize(rowCount - 1).Value = Application.WorksheetFunction.Transpose(outputDataA)
    ThisWorkbook.Sheets(targetSheetName).Range("B1").Resize(rowCount - 1).Value = Application.WorksheetFunction.Transpose(outputDataB)

    ' Unhide lembar kerja tujuan
    ThisWorkbook.Sheets(targetSheetName).Visible = xlSheetVisible
End Sub

