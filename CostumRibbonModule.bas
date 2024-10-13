' Untuk menentukan apakah tab tertentu aktif atau tidak
Sub GetEnabled(control As IRibbonControl, ByRef MakeVisible)
    Dim SheetName As String
    SheetName = "DEV"
    On Error Resume Next
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets(SheetName)
    Select Case control.ID
        Case "ApplicationOptionsDialog":        MakeVisible = ws.Range("L3").value
        Case "TabInfo":                         MakeVisible = ws.Range("L4").value
        Case "TabOfficeStart":                  MakeVisible = ws.Range("L5").value
        Case "TabRecent":                       MakeVisible = ws.Range("L6").value
        Case "TabSave":                         MakeVisible = ws.Range("L7").value
        Case "TabPrint":                        MakeVisible = ws.Range("L8").value
        Case "ShareDocument":                   MakeVisible = ws.Range("L9").value
        Case "Publish2Tab":                     MakeVisible = ws.Range("L10").value
        Case "TabPublish":                      MakeVisible = ws.Range("L11").value
        Case "TabHelp":                         MakeVisible = ws.Range("L12").value
        Case "TabOfficeFeedback":               MakeVisible = ws.Range("L13").value
        Case "FileSave":                        MakeVisible = ws.Range("L14").value
        Case "HistoryTab":                      MakeVisible = ws.Range("L15").value
        Case "FileClose":                       MakeVisible = ws.Range("L16").value
        Case Else:                              MakeVisible = False
    End Select
End Sub

' Subroutine untuk menentukan visibilitas tab berdasarkan nilai variabel
Sub GetVisible(control As IRibbonControl, ByRef MakeVisible)
    Dim SheetName As String
    SheetName = "DEV"
    On Error Resume Next
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets(SheetName)
    Select Case control.ID
        Case "TabHome":                 MakeVisible = ws.Range("L17").value
        Case "TabView":                 MakeVisible = ws.Range("L18").value
        Case "TabReview":               MakeVisible = ws.Range("L19").value
        Case "TabData":                 MakeVisible = ws.Range("L20").value
        Case "TabAutomate":             MakeVisible = ws.Range("L21").value
        Case "TabInsert":               MakeVisible = ws.Range("L22").value
        Case "TabPageLayoutExcel":      MakeVisible = ws.Range("L23").value
        Case "TabAddIns":               MakeVisible = ws.Range("L24").value
        Case "TabFormulas":             MakeVisible = ws.Range("L25").value
        Case "TabDeveloper":            MakeVisible = ws.Range("L26").value
        Case Else:                      MakeVisible = False
    End Select
End Sub
