Function SumByColor(CellColor As Range, SumRange As Range, EndColor As Range) As Double
    Dim Color As Long
    Dim EndColorValue As Long
    Dim Total As Double
    Dim cell As Range
    
    Color = CellColor.Interior.Color
    EndColorValue = EndColor.Interior.Color
    
    For Each cell In SumRange
        If cell.Interior.Color = EndColorValue Then
            Exit For
        End If
        If cell.Interior.Color = Color Then
            Total = Total + cell.Value
        End If
    Next cell
    
    SumByColor = Total
End Function

