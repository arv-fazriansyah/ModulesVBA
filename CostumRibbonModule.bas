' Untuk menentukan apakah tab tertentu aktif atau tidak
Sub GetEnabled(control As IRibbonControl, ByRef MakeVisible)
    Dim SheetName As String
    SheetName = "DEV"
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets(SheetName)
    Select Case control.ID
        Case "ApplicationOptionsDialog":        MakeVisible = ws.Range("I3").value
        Case "TabInfo":                         MakeVisible = ws.Range("I4").value
        Case "TabOfficeStart":                  MakeVisible = ws.Range("I5").value
        Case "TabRecent":                       MakeVisible = ws.Range("I6").value
        Case "TabSave":                         MakeVisible = ws.Range("I7").value
        Case "TabPrint":                        MakeVisible = ws.Range("I8").value
        Case "ShareDocument":                   MakeVisible = ws.Range("I9").value
        Case "Publish2Tab":                     MakeVisible = ws.Range("I10").value
        Case "TabPublish":                      MakeVisible = ws.Range("I11").value
        Case "TabHelp":                         MakeVisible = ws.Range("I12").value
        Case "TabOfficeFeedback":               MakeVisible = ws.Range("I13").value
        Case "FileSave":                        MakeVisible = ws.Range("I14").value
        Case "HistoryTab":                      MakeVisible = ws.Range("I15").value
        Case "FileClose":                       MakeVisible = ws.Range("I16").value
        Case Else:                              MakeVisible = False
    End Select
End Sub

' Subroutine untuk menentukan visibilitas tab berdasarkan nilai variabel
Sub GetVisible(control As IRibbonControl, ByRef MakeVisible)
    Dim SheetName As String
    SheetName = "DEV"
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets(SheetName)
    Select Case control.ID
        Case "TabHome":                 MakeVisible = ws.Range("I17").value
        Case "TabView":                 MakeVisible = ws.Range("I18").value
        Case "TabReview":               MakeVisible = ws.Range("I19").value
        Case "TabData":                 MakeVisible = ws.Range("I20").value
        Case "TabAutomate":             MakeVisible = ws.Range("I21").value
        Case "TabInsert":               MakeVisible = ws.Range("I22").value
        Case "TabPageLayoutExcel":      MakeVisible = ws.Range("I23").value
        Case "TabAddIns":               MakeVisible = ws.Range("I24").value
        Case "TabFormulas":             MakeVisible = ws.Range("I25").value
        Case "TabDeveloper":            MakeVisible = ws.Range("I26").value
        Case Else:                      MakeVisible = False
    End Select
End Sub

