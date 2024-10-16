Sub Dashboard(ByVal control As IRibbonControl)
Unhide.Menu
End Sub
Sub Update(ByVal control As IRibbonControl)
BtnUpdate.GsheetDataUpdate
End Sub
Sub Upload(ByVal control As IRibbonControl)
UploadFile.UploadFile1
End Sub
Sub PrintView(ByVal control As IRibbonControl)
Dev.PrintActiveSheet
End Sub
Sub Saved(ByVal control As IRibbonControl)
Dev.Simpan
End Sub
Sub PetaBenahi(ByVal control As IRibbonControl)
Unhide.Peta_Benahi
End Sub
Sub LembarRKT(ByVal control As IRibbonControl)
Unhide.Lembar_RKT
End Sub
Sub LembarRKAS(ByVal control As IRibbonControl)
Unhide.Lembar_RKAS
End Sub
Sub data(ByVal control As IRibbonControl)
Unhide.DataAwal
End Sub
Sub Matrix(ByVal control As IRibbonControl)
Unhide.DataMatrix
End Sub
Sub HarsatBarjas(ByVal control As IRibbonControl)
Unhide.DataHarsatBarjas
End Sub
Sub HarsatModal(ByVal control As IRibbonControl)
Unhide.DataHarsatModal
End Sub
Sub RKASROB(ByVal control As IRibbonControl)
Unhide.RKAS_ROB
End Sub
Sub RKASPerTahap(ByVal control As IRibbonControl)
Unhide.RKAS_TAHAP
End Sub
Sub RKASSNP(ByVal control As IRibbonControl)
Unhide.RKAS_SNP
End Sub
Sub RKASSIPD(ByVal control As IRibbonControl)
Unhide.RKAS_SIPD
End Sub
Sub RBK(ByVal control As IRibbonControl)
Unhide.RBK_1
End Sub
Sub RBK2(ByVal control As IRibbonControl)
Unhide.RBK_2
End Sub
Sub RBKReload(ByVal control As IRibbonControl)
DuplikatRBK.DuplicateSheet
End Sub
Sub AnalisisGugus(ByVal control As IRibbonControl)
Unhide.AnGugus
End Sub
Sub AnalisisBuku(ByVal control As IRibbonControl)
Unhide.AnBuku
End Sub
Sub AnalisisEkskul(ByVal control As IRibbonControl)
Unhide.AnEkskul
End Sub
Sub AnalisisHonor(ByVal control As IRibbonControl)
Unhide.AnHonor
End Sub
Sub CoverRKAS(ByVal control As IRibbonControl)
Download.DownCover
End Sub
Sub CoverRKASPerubahan(ByVal control As IRibbonControl)
Download.DownCoverRKAS
End Sub
Sub SKBendahara(ByVal control As IRibbonControl)
Download.DownSKBendahara
End Sub
Sub SKTimBOS(ByVal control As IRibbonControl)
Download.DownSKTimBOS
End Sub
Sub SKTimPBJSekolah(ByVal control As IRibbonControl)
Download.DownSKTimPBJ
End Sub
Sub BeritaAcara(ByVal control As IRibbonControl)
Download.DownBeritaAcara
End Sub
Sub LembarPengesahan(ByVal control As IRibbonControl)
Download.DownLembarPengesahan
End Sub
Sub ConvertPDF(ByVal control As IRibbonControl)
Convert2PDF.ConvertToPDF
End Sub

' Untuk menentukan apakah tab tertentu aktif atau tidak
Sub GetEnabled(control As IRibbonControl, ByRef MakeVisible)
    Dim sheetName As String
    sheetName = "DEV"
    On Error Resume Next
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets(sheetName)
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
    Dim sheetName As String
    sheetName = "DEV"
    On Error Resume Next
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets(sheetName)
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

        Case "customTab":               MakeVisible = ws.Range("L27").value
        Case "customGroup1":            MakeVisible = ws.Range("L28").value
        Case "customGroup2":            MakeVisible = ws.Range("L29").value
        Case "customGroup3":            MakeVisible = ws.Range("L30").value
        Case "customGroup4":            MakeVisible = ws.Range("L31").value
        Case "customGroup5":            MakeVisible = ws.Range("L32").value
        Case "customGroup6":            MakeVisible = ws.Range("L33").value
        
        Case "Dash":                    MakeVisible = ws.Range("L34").value
        Case "Update":                  MakeVisible = ws.Range("L35").value
        Case "Upload":                  MakeVisible = ws.Range("L36").value
        Case "PetaBenahi":              MakeVisible = ws.Range("L37").value
        Case "LembarRKT":               MakeVisible = ws.Range("L38").value
        Case "LembarRKAS":              MakeVisible = ws.Range("L39").value
        Case "PrintView":               MakeVisible = ws.Range("L40").value
        Case "Saved":                   MakeVisible = ws.Range("L41").value
        
        Case "Data":                    MakeVisible = ws.Range("L42").value
        Case "Matrix":                  MakeVisible = ws.Range("L43").value
        Case "HarsatBarjas":            MakeVisible = ws.Range("L44").value
        Case "HarsatModal":             MakeVisible = ws.Range("L45").value
        
        Case "AnalisisGugus":           MakeVisible = ws.Range("L46").value
        Case "AnalisisBuku":            MakeVisible = ws.Range("L47").value
        Case "AnalisisEkskul":          MakeVisible = ws.Range("L48").value
        Case "AnalisisHonor":           MakeVisible = ws.Range("L49").value
        
        Case "RKASROB":                 MakeVisible = ws.Range("L50").value
        Case "RKASPerTahap":            MakeVisible = ws.Range("L51").value
        Case "RKASSNP":                 MakeVisible = ws.Range("L52").value
        Case "RKASSIPD":                MakeVisible = ws.Range("L53").value
        
        Case "RBK":                     MakeVisible = ws.Range("L54").value
        Case "RBK2":                    MakeVisible = ws.Range("L55").value
        Case "RBKReload":               MakeVisible = ws.Range("L56").value
        
        Case "CoverRKAS":               MakeVisible = ws.Range("L57").value
        Case "CoverRKASPerubahan":      MakeVisible = ws.Range("L58").value
        Case "SKBendahara":             MakeVisible = ws.Range("L59").value
        Case "SKTimBOS":                MakeVisible = ws.Range("L60").value
        Case "SKTimPBJSekolah":         MakeVisible = ws.Range("L61").value
        Case "BeritaAcara":             MakeVisible = ws.Range("L62").value
        Case "LembarPengesahan":        MakeVisible = ws.Range("L63").value
        Case "ConvertPDF":              MakeVisible = ws.Range("L64").value
        Case Else:                      MakeVisible = False
    End Select
End Sub
