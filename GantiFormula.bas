Sub CopyFormulas()
    Dim sourceSheetName As String
    Dim sourceColumnFormula As String
    Dim destinationColumnSheet As String
    Dim destinationColumnCell As String
    
    sourceSheetName = "DATAUSER"
    sourceColumnFormula = "H"
    destinationColumnSheet = "I"
    destinationColumnCell = "J"
    
    Dim sourceSheet As Worksheet
    Set sourceSheet = ThisWorkbook.Sheets(sourceSheetName)
    
    Dim separator As String
    ' Get the list separator from the user's regional settings
    separator = Application.International(xlListSeparator)
    
    Dim lastRow As Long
    lastRow = sourceSheet.Cells(sourceSheet.Rows.count, sourceColumnFormula).End(xlUp).row
    
    Dim i As Long
    For i = 1 To lastRow
        ' Get the formula from column H
        Dim formulaValue As String
        formulaValue = sourceSheet.Cells(i, sourceColumnFormula).formula
        
        ' Replace the formula separator with the regional separator
        formulaValue = Replace(formulaValue, ";", separator)
        formulaValue = Replace(formulaValue, ",", separator)
        
        ' Get the destination sheet name from column I
        Dim destSheetName As String
        destSheetName = sourceSheet.Cells(i, destinationColumnSheet).value
        
        ' Get the destination cell from column J
        Dim destCell As String
        destCell = sourceSheet.Cells(i, destinationColumnCell).value
        
        ' Check if the destination sheet name and cell are not empty
        If destSheetName <> "" And destCell <> "" Then
            Dim destSheet As Worksheet
            On Error Resume Next
            Set destSheet = ThisWorkbook.Sheets(destSheetName)
            On Error GoTo 0
            
            If Not destSheet Is Nothing Then
                ' Paste the value as text into the destination cell on the destination sheet
                Application.DisplayAlerts = False
                destSheet.Range(destCell).value = formulaValue
                Application.DisplayAlerts = True
            End If
        End If
    Next i
    
    ' Delete external workbook links
    Dim links As Variant
    links = ThisWorkbook.LinkSources(xlExcelLinks)
    
    If Not IsEmpty(links) Then
        For i = 1 To UBound(links)
            ThisWorkbook.BreakLink Name:=links(i), Type:=xlLinkTypeExcelLinks
        Next i
    End If
End Sub
