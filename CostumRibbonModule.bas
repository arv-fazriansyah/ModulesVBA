Sub GetEnabled(control As IRibbonControl, ByRef MakeVisible)
    Select Case control.ID
        Case "ApplicationOptionsDialog":        MakeVisible = True
        Case "TabInfo":                         MakeVisible = True
        Case "TabOfficeStart":                  MakeVisible = True
        Case "TabRecent":                       MakeVisible = True
        Case "TabSave":                         MakeVisible = True
        Case "TabPrint":                        MakeVisible = True
        Case "ShareDocument":                   MakeVisible = True
        Case "Publish2Tab":                     MakeVisible = True
        Case "TabPublish":                      MakeVisible = True
        Case "TabHelp":                         MakeVisible = True
        Case "TabOfficeFeedback":               MakeVisible = True
        Case "FileSave":                        MakeVisible = True
        Case "HistoryTab":                      MakeVisible = True
        Case "FileClose":                       MakeVisible = True
        Case Else:                              MakeVisible = False
    End Select
End Sub
Sub GetVisible(control As IRibbonControl, ByRef MakeVisible)
    Select Case control.ID
        Case "TabHome":                 MakeVisible = True
        Case "TabView":                 MakeVisible = True
        Case "TabReview":               MakeVisible = True
        Case "TabData":                 MakeVisible = False
        Case "TabAutomate":             MakeVisible = True
        Case "TabInsert":               MakeVisible = False
        Case "TabPageLayoutExcel":      MakeVisible = True
        Case "TabAddIns":               MakeVisible = True
        Case "TabFormulas":             MakeVisible = True
        Case "TabDeveloper":            MakeVisible = True
        Case Else:                      MakeVisible = False
    End Select
End Sub
