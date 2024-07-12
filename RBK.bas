Sub ApplyFormulasBasedOnColors()
    Dim ws As Worksheet
    Dim cell As Range
    Dim colorValue As Long
    Dim formulasF As Variant
    Dim formulasOtherCols As Variant
    Dim columns As Variant
    Dim i As Long
    Dim startRow As Long
    Dim lastRow As Long
    Dim formula As String
    
    ' Ganti "RBK" dengan nama sheet yang sesuai
    Set ws = ThisWorkbook.Sheets("RBK")
    
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    
    ' Mendapatkan baris terakhir yang digunakan di kolom E
    lastRow = ws.Cells(ws.Rows.Count, "E").End(xlUp).row
    
    ' Inisialisasi array formulas
    formulasF = ws.Range("F17:F" & lastRow).formula
    columns = Array("N", "V", "AD", "AL", "AT", "BB", "BJ", "BR", "BZ", "CH", "CP", "CX")
    ReDim formulasOtherCols(LBound(columns) To UBound(columns))
    
    ' Loop untuk kolom F
    For Each cell In ws.Range("F17:F" & lastRow)
        If ws.Cells(cell.row, "E").Value = "" Then
            formulasF(cell.row - 16, 1) = ""
        Else
            colorValue = cell.Interior.color
            Select Case colorValue
                Case RGB(255, 255, 0), RGB(102, 204, 255), RGB(255, 217, 102)
                    formulasF(cell.row - 16, 1) = "=SUM(G" & cell.row & ":CX" & cell.row & ")"
                Case RGB(255, 255, 255)
                    formula = "=SUM("
                    For i = LBound(columns) To UBound(columns)
                        formula = formula & columns(i) & cell.row & ","
                    Next i
                    formula = Left(formula, Len(formula) - 1) & ")"
                    formulasF(cell.row - 16, 1) = formula
            End Select
        End If
    Next cell

    ' Loop untuk kolom lainnya
    For i = LBound(columns) To UBound(columns)
        formulasOtherCols(i) = ws.Range(columns(i) & "17:" & columns(i) & lastRow).formula
        startRow = 18
        
        For Each cell In ws.Range(columns(i) & "17:" & columns(i) & lastRow)
            If ws.Cells(cell.row, "E").Value = "" Then
                formulasOtherCols(i)(cell.row - 16, 1) = ""
            Else
                colorValue = cell.Interior.color
                Select Case colorValue
                    Case RGB(255, 255, 0), RGB(102, 204, 255), RGB(255, 217, 102)
                        formula = "=SumByColor(" & columns(i) & startRow & "," & columns(i) & startRow & ":" & columns(i) & lastRow & "," & columns(i) & (startRow - 1) & ")"
                        formulasOtherCols(i)(cell.row - 16, 1) = formula
                    Case RGB(255, 255, 255)
                        Select Case columns(i)
                            Case "N"
                                formula = "=SUM(G" & cell.row & "*I" & cell.row & "*K" & cell.row & "*M" & cell.row & ")"
                            Case "V"
                                formula = "=SUM(O" & cell.row & "*Q" & cell.row & "*S" & cell.row & "*U" & cell.row & ")"
                            Case "AD"
                                formula = "=SUM(W" & cell.row & "*Y" & cell.row & "*AA" & cell.row & "*AC" & cell.row & ")"
                            Case "AL"
                                formula = "=SUM(AE" & cell.row & "*AG" & cell.row & "*AI" & cell.row & "*AK" & cell.row & ")"
                            Case "AT"
                                formula = "=SUM(AM" & cell.row & "*AO" & cell.row & "*AQ" & cell.row & "*AS" & cell.row & ")"
                            Case "BB"
                                formula = "=SUM(AU" & cell.row & "*AW" & cell.row & "*AY" & cell.row & "*BA" & cell.row & ")"
                            Case "BJ"
                                formula = "=SUM(BC" & cell.row & "*BE" & cell.row & "*BG" & cell.row & "*BI" & cell.row & ")"
                            Case "BR"
                                formula = "=SUM(BK" & cell.row & "*BM" & cell.row & "*BO" & cell.row & "*BQ" & cell.row & ")"
                            Case "BZ"
                                formula = "=SUM(BS" & cell.row & "*BU" & cell.row & "*BW" & cell.row & "*BY" & cell.row & ")"
                            Case "CH"
                                formula = "=SUM(CA" & cell.row & "*CC" & cell.row & "*CE" & cell.row & "*CG" & cell.row & ")"
                            Case "CP"
                                formula = "=SUM(CI" & cell.row & "*CK" & cell.row & "*CM" & cell.row & "*CO" & cell.row & ")"
                            Case "CX"
                                formula = "=SUM(CQ" & cell.row & "*CS" & cell.row & "*CU" & cell.row & "*CW" & cell.row & ")"
                        End Select
                        formulasOtherCols(i)(cell.row - 16, 1) = formula
                End Select
            End If
            
            startRow = startRow + 1
        Next cell
    Next i

    ' Menulis kembali array formulas ke worksheet
    ws.Range("F17:F" & lastRow).formula = formulasF
    For i = LBound(columns) To UBound(columns)
        ws.Range(columns(i) & "17:" & columns(i) & lastRow).formula = formulasOtherCols(i)
    Next i

    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
End Sub

