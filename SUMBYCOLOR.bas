Function SumByColor(CellColor As Range, SumRange As Range)
    Dim Color As Long
    Dim Total As Double
    Dim Cell As Range
    
    Color = CellColor.Interior.Color
    
    For Each Cell In SumRange
        If Cell.Interior.Color = Color Then
            Total = Total + Cell.Value
        End If
    Next Cell
    
    SumByColor = Total
End Function
