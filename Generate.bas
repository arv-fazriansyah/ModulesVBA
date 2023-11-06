Sub ExtractData()
    Dim dataColumnA As Range
    Dim dataColumnB As Range
    Dim dataColumnC As Range
    Dim dataColumnD As Range ' New data column
    Dim dataColumnE As Range ' New data column
    Dim cellA As Range
    Dim cellB As Range
    Dim cellC As Range
    Dim cellD As Range ' New data column
    Dim cellE As Range ' New data column
    Dim Ws As Worksheet
    Dim outputRangeA As Range
    Dim outputRangeB As Range
    Dim outputRangeC As Range
    Dim outputRangeD As Range ' New output column
    Dim outputRangeE As Range ' New output column
    Dim result As String
    Dim count As Integer
    Dim outputDataA() As Variant
    Dim outputDataB() As Variant
    Dim outputDataC() As Variant
    Dim outputDataD() As Variant ' New output array
    Dim outputDataE() As Variant ' New output array
    Dim rowCount As Integer

    Set Ws = ThisWorkbook.Sheets("RAPOR1")
    Set dataColumnA = Ws.Range("B6:B200")
    Set dataColumnB = Ws.Range("G6:G200")
    Set dataColumnC = Ws.Range("I1:I200")
    Set dataColumnD = Ws.Range("C6:C200") ' New data column
    Set dataColumnE = Ws.Range("D6:D200") ' New data column

    Dim targetSheetName As String
    targetSheetName = "DATARAPOR1"
    On Error Resume Next
    Set outputRangeA = ThisWorkbook.Sheets(targetSheetName).Range("A1")
    Set outputRangeB = ThisWorkbook.Sheets(targetSheetName).Range("B1")
    Set outputRangeC = ThisWorkbook.Sheets(targetSheetName).Range("C1")
    Set outputRangeD = ThisWorkbook.Sheets(targetSheetName).Range("D1") ' New output column
    Set outputRangeE = ThisWorkbook.Sheets(targetSheetName).Range("E1") ' New output column
    On Error GoTo 0

    If outputRangeA Is Nothing Then
        ThisWorkbook.Sheets.Add(, ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count)).Name = targetSheetName
        Set outputRangeA = ThisWorkbook.Sheets(targetSheetName).Range("A1")
        Set outputRangeB = ThisWorkbook.Sheets(targetSheetName).Range("B1")
        Set outputRangeC = ThisWorkbook.Sheets(targetSheetName).Range("C1")
        Set outputRangeD = ThisWorkbook.Sheets(targetSheetName).Range("D1") ' New output column
        Set outputRangeE = ThisWorkbook.Sheets(targetSheetName).Range("E1") ' New output column
    End If

    rowCount = 1

    For Each cellA In dataColumnA.Cells
        dataA = cellA.Value
        ReDim Preserve outputDataA(1 To rowCount)
        dataB = dataColumnB.Cells(cellA.Row - 5, 1).Value ' Kolom B dimulai dari baris ke-6
        ReDim Preserve outputDataB(1 To rowCount)
        dataC = dataColumnC.Cells(cellA.Row, 1).Value
        ReDim Preserve outputDataC(1 To rowCount)
        dataD = dataColumnD.Cells(cellA.Row - 5, 1).Value ' New data column
        ReDim Preserve outputDataD(1 To rowCount) ' New output array
        dataE = dataColumnE.Cells(cellA.Row - 5, 1).Value ' New data column
        ReDim Preserve outputDataE(1 To rowCount) ' New output array

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
        outputDataD(rowCount) = dataD ' New output array
        outputDataE(rowCount) = dataE ' New output array

        rowCount = rowCount + 1

        For j = LBound(sentences) To UBound(sentences)
            outputRangeA.Value = outputDataA(rowCount - 1)
            outputRangeB.Value = outputDataB(rowCount - 1)
            outputRangeC.Value = sentences(j)
            outputRangeD.Value = outputDataD(rowCount - 1) ' New output column
            outputRangeE.Value = outputDataE(rowCount - 1) ' New output column
            If j < UBound(sentences) Then
                Set outputRangeA = outputRangeA.Offset(1, 0)
                Set outputRangeB = outputRangeB.Offset(1, 0)
                Set outputRangeC = outputRangeC.Offset(1, 0)
                Set outputRangeD = outputRangeD.Offset(1, 0) ' New output column
                Set outputRangeE = outputRangeE.Offset(1, 0) ' New output column
            End If
        Next j

        If rowCount > 1 And cellA.Offset(1, 0).Value <> "" Then
            Set outputRangeA = outputRangeA.Offset(1, 0)
            Set outputRangeB = outputRangeB.Offset(1, 0)
            Set outputRangeC = outputRangeC.Offset(1, 0)
            Set outputRangeD = outputRangeD.Offset(1, 0) ' New output column
            Set outputRangeE = outputRangeE.Offset(1, 0) ' New output column
            rowCount = rowCount + 1
        End If
    Next cellA

    ThisWorkbook.Sheets(targetSheetName).Range("A1").Resize(rowCount - 1).Value = Application.WorksheetFunction.Transpose(outputDataA)
    ThisWorkbook.Sheets(targetSheetName).Range("B1").Resize(rowCount - 1).Value = Application.WorksheetFunction.Transpose(outputDataB)
    ThisWorkbook.Sheets(targetSheetName).Range("D1").Resize(rowCount - 1).Value = Application.WorksheetFunction.Transpose(outputDataD) ' New output column
    ThisWorkbook.Sheets(targetSheetName).Range("E1").Resize(rowCount - 1).Value = Application.WorksheetFunction.Transpose(outputDataE) ' New output column

    ' Unhide lembar kerja tujuan
    ThisWorkbook.Sheets(targetSheetName).Visible = xlSheetVisible
End Sub
