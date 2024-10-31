Sub Dashboard(ByVal control As IRibbonControl)
On Error Resume Next
Unhide.Menu
End Sub
Sub Update(ByVal control As IRibbonControl)
On Error Resume Next
BtnUpdate.DataUpdate
End Sub
Sub Upload(ByVal control As IRibbonControl)
On Error Resume Next
UploadFile.UploadFile1
End Sub
Sub PrintView(ByVal control As IRibbonControl)
On Error Resume Next
Dev.PrintActiveSheet
End Sub
Sub Saved(ByVal control As IRibbonControl)
On Error Resume Next
Dev.Simpan
End Sub
Sub PetaBenahi(ByVal control As IRibbonControl)
On Error Resume Next
Unhide.Peta_Benahi
End Sub
Sub LembarRKT(ByVal control As IRibbonControl)
On Error Resume Next
Unhide.Lembar_RKT
End Sub
Sub LembarRKAS(ByVal control As IRibbonControl)
On Error Resume Next
Unhide.Lembar_RKAS
End Sub
Sub Data(ByVal control As IRibbonControl)
On Error Resume Next
Unhide.DataAwal
End Sub
Sub DataRapat(ByVal control As IRibbonControl)
On Error Resume Next
Unhide.DataRapats
End Sub
Sub Matrix(ByVal control As IRibbonControl)
On Error Resume Next
Unhide.DataMatrix
End Sub
Sub HarsatBarjas(ByVal control As IRibbonControl)
On Error Resume Next
Unhide.DataHarsatBarjas
End Sub
Sub HarsatModal(ByVal control As IRibbonControl)
On Error Resume Next
Unhide.DataHarsatModal
End Sub
Sub RKASROB(ByVal control As IRibbonControl)
On Error Resume Next
Unhide.RKAS_ROB
End Sub
Sub RKASPerTahap(ByVal control As IRibbonControl)
On Error Resume Next
Unhide.RKAS_TAHAP
End Sub
Sub RKASSNP(ByVal control As IRibbonControl)
On Error Resume Next
Unhide.RKAS_SNP
End Sub
Sub RKASSIPD(ByVal control As IRibbonControl)
On Error Resume Next
Unhide.RKAS_SIPD
End Sub
Sub KomponenBOS(ByVal control As IRibbonControl)
On Error Resume Next
Unhide.Komponen_BOS
End Sub
Sub RBK(ByVal control As IRibbonControl)
On Error Resume Next
Unhide.RBK_1
End Sub
Sub Planning1(ByVal control As IRibbonControl)
On Error Resume Next

End Sub
Sub Planning2(ByVal control As IRibbonControl)
On Error Resume Next

End Sub
Sub PlanningTahun(ByVal control As IRibbonControl)
On Error Resume Next

End Sub
Sub AnalisisGugus(ByVal control As IRibbonControl)
On Error Resume Next
Unhide.AnGugus
End Sub
Sub AnalisisBuku(ByVal control As IRibbonControl)
On Error Resume Next
Unhide.AnBuku
End Sub
Sub AnalisisEkskul(ByVal control As IRibbonControl)
On Error Resume Next
Unhide.AnEkskul
End Sub
Sub AnalisisHonor(ByVal control As IRibbonControl)
On Error Resume Next
Unhide.AnHonor
End Sub
Sub CoverRKAS(ByVal control As IRibbonControl)
On Error Resume Next
Download.DownCover
End Sub
Sub CoverRKASPerubahan(ByVal control As IRibbonControl)
On Error Resume Next
Download.DownCoverRKAS
End Sub
Sub SKBendahara(ByVal control As IRibbonControl)
On Error Resume Next
Download.DownSKBendahara
End Sub
Sub SKTimBOS(ByVal control As IRibbonControl)
On Error Resume Next
Download.DownSKTimBOS
End Sub
Sub SKTimPBJSekolah(ByVal control As IRibbonControl)
On Error Resume Next
Download.DownSKTimPBJ
End Sub
Sub BeritaAcara(ByVal control As IRibbonControl)
On Error Resume Next
Download.DownBeritaAcara
End Sub
Sub LembarPengesahan(ByVal control As IRibbonControl)
On Error Resume Next
Download.DownLembarPengesahan
End Sub
Sub ConvertPDF(ByVal control As IRibbonControl)
On Error Resume Next
Convert2PDF.ConvertToPDF
End Sub

' Untuk menentukan apakah tab tertentu aktif atau tidak
Sub GetEnabled(control As IRibbonControl, ByRef MakeVisible)
    Dim sheetName As String
    sheetName = "DEV"
    On Error Resume Next
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets(sheetName)
    Select Case control.Id
        Case "ApplicationOptionsDialog":        MakeVisible = ws.range("L3").value
        Case "TabInfo":                         MakeVisible = ws.range("L4").value
        Case "TabOfficeStart":                  MakeVisible = ws.range("L5").value
        Case "TabRecent":                       MakeVisible = ws.range("L6").value
        Case "TabSave":                         MakeVisible = ws.range("L7").value
        Case "TabPrint":                        MakeVisible = ws.range("L8").value
        Case "ShareDocument":                   MakeVisible = ws.range("L9").value
        Case "Publish2Tab":                     MakeVisible = ws.range("L10").value
        Case "TabPublish":                      MakeVisible = ws.range("L11").value
        Case "TabHelp":                         MakeVisible = ws.range("L12").value
        Case "TabOfficeFeedback":               MakeVisible = ws.range("L13").value
        Case "FileSave":                        MakeVisible = ws.range("L14").value
        Case "HistoryTab":                      MakeVisible = ws.range("L15").value
        Case "FileClose":                       MakeVisible = ws.range("L16").value
        Case Else:                              MakeVisible = False
    End Select
End Sub

' Subroutine untuk menentukan visibilitas tab berdasarkan nilai variabel
Sub GetVisible(control As IRibbonControl, ByRef MakeVisible)
    Dim sheetName As String
    sheetName = "DEV"
    On Error Resume Next
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets(sheetName)
    Select Case control.Id
        Case "TabHome":                 MakeVisible = ws.range("L17").value
        Case "TabView":                 MakeVisible = ws.range("L18").value
        Case "TabReview":               MakeVisible = ws.range("L19").value
        Case "TabData":                 MakeVisible = ws.range("L20").value
        Case "TabAutomate":             MakeVisible = ws.range("L21").value
        Case "TabInsert":               MakeVisible = ws.range("L22").value
        Case "TabPageLayoutExcel":      MakeVisible = ws.range("L23").value
        Case "TabAddIns":               MakeVisible = ws.range("L24").value
        Case "TabFormulas":             MakeVisible = ws.range("L25").value
        Case "TabDeveloper":            MakeVisible = ws.range("L26").value

        Case "customTab":               MakeVisible = ws.range("L27").value
        Case "customGroup1":            MakeVisible = ws.range("L28").value
        Case "customGroup2":            MakeVisible = ws.range("L29").value
        Case "customGroup3":            MakeVisible = ws.range("L30").value
        Case "customGroup4":            MakeVisible = ws.range("L31").value
        Case "customGroup5":            MakeVisible = ws.range("L32").value
        Case "customGroup6":            MakeVisible = ws.range("L33").value
        Case "customGroup7":            MakeVisible = ws.range("L34").value
        
        Case "Dash":                    MakeVisible = ws.range("L35").value
        Case "Update":                  MakeVisible = ws.range("L36").value
        Case "Upload":                  MakeVisible = ws.range("L37").value
        Case "PetaBenahi":              MakeVisible = ws.range("L38").value
        Case "LembarRKT":               MakeVisible = ws.range("L39").value
        Case "LembarRKAS":              MakeVisible = ws.range("L40").value
        Case "PrintView":               MakeVisible = ws.range("L41").value
        Case "Saved":                   MakeVisible = ws.range("L42").value
        
        Case "Data":                    MakeVisible = ws.range("L43").value
        Case "DataRapat":               MakeVisible = ws.range("L44").value
        Case "Matrix":                  MakeVisible = ws.range("L45").value
        Case "HarsatBarjas":            MakeVisible = ws.range("L46").value
        Case "HarsatModal":             MakeVisible = ws.range("L47").value
        
        Case "AnalisisGugus":           MakeVisible = ws.range("L48").value
        Case "AnalisisBuku":            MakeVisible = ws.range("L49").value
        Case "AnalisisEkskul":          MakeVisible = ws.range("L50").value
        Case "AnalisisHonor":           MakeVisible = ws.range("L51").value
        
        Case "RKASROB":                 MakeVisible = ws.range("L52").value
        Case "RKASPerTahap":            MakeVisible = ws.range("L53").value
        Case "RKASSNP":                 MakeVisible = ws.range("L54").value
        Case "RKASSIPD":                MakeVisible = ws.range("L55").value
        Case "KomponenBOS":             MakeVisible = ws.range("L56").value
        
        Case "RBK":                     MakeVisible = ws.range("L57").value
        Case "Planning1":               MakeVisible = ws.range("L58").value
        Case "Planning2":               MakeVisible = ws.range("L59").value
        Case "PlanningTahun":           MakeVisible = ws.range("L60").value
        
        Case "CoverRKAS":               MakeVisible = ws.range("L61").value
        Case "CoverRKASPerubahan":      MakeVisible = ws.range("L62").value
        Case "SKBendahara":             MakeVisible = ws.range("L63").value
        Case "SKTimBOS":                MakeVisible = ws.range("L64").value
        Case "SKTimPBJSekolah":         MakeVisible = ws.range("L65").value
        Case "BeritaAcara":             MakeVisible = ws.range("L66").value
        Case "LembarPengesahan":        MakeVisible = ws.range("L67").value
        Case "Verval":                  MakeVisible = ws.range("L68").value
        Case Else:                      MakeVisible = False
    End Select
End Sub
