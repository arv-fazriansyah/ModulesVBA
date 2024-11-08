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

Sub RBK(ByVal control As IRibbonControl)
On Error Resume Next
Unhide.RBK_1
End Sub
Sub ReloadRBK(ByVal control As IRibbonControl)
On Error Resume Next
ForRBK.SumColor1
End Sub

Sub Planning1(ByVal control As IRibbonControl)
On Error Resume Next
DuplikatRBK.semester1
End Sub
Sub Planning2(ByVal control As IRibbonControl)
On Error Resume Next
DuplikatRBK.semester2
End Sub
Sub PlanningTahun(ByVal control As IRibbonControl)
On Error Resume Next
DuplikatRBK.setahun
End Sub

Sub RKASPerTahap(ByVal control As IRibbonControl)
On Error Resume Next
Unhide.RKAS_TAHAP
End Sub
Sub RKASROB(ByVal control As IRibbonControl)
On Error Resume Next
Unhide.RKAS_ROB
End Sub
Sub RKASSIPD(ByVal control As IRibbonControl)
On Error Resume Next
Unhide.RKAS_SIPD
End Sub
Sub RKASSNP(ByVal control As IRibbonControl)
On Error Resume Next
Unhide.RKAS_SNP
End Sub
Sub KomponenBOS(ByVal control As IRibbonControl)
On Error Resume Next
Unhide.Komponen_BOS
End Sub
Sub RekonSaldo(ByVal control As IRibbonControl)
On Error Resume Next
Unhide.Rekon_Saldo
End Sub

Sub LembarPengesahan(ByVal control As IRibbonControl)
On Error Resume Next
Download.DownLembarPengesahan
End Sub

Sub PenyusunanRKAS(ByVal control As IRibbonControl)
On Error Resume Next
Download.DownPenyusunanRKAS
End Sub
Sub BelanjaModal(ByVal control As IRibbonControl)
On Error Resume Next
Download.DownBelanjaModal
End Sub
Sub PenggunaanDana(ByVal control As IRibbonControl)
On Error Resume Next
Download.DownPenggunaanDana
End Sub

Sub CoverRKAS(ByVal control As IRibbonControl)
On Error Resume Next
Download.DownCoverRKAS
End Sub
Sub CoverRKASPerubahan(ByVal control As IRibbonControl)
On Error Resume Next
Download.DownCoverRKASPerubahan
End Sub

Sub SKTimBOS(ByVal control As IRibbonControl)
On Error Resume Next
Download.DownSKTimBOS
End Sub
Sub SKTimPBJSekolah(ByVal control As IRibbonControl)
On Error Resume Next
Download.DownSKTimPBJ
End Sub
Sub SKBendahara(ByVal control As IRibbonControl)
On Error Resume Next
Download.DownSKBendahara
End Sub
Sub SKTAS(ByVal control As IRibbonControl)
On Error Resume Next
Download.DownSKTAS
End Sub

Sub Verval(ByVal control As IRibbonControl)
On Error Resume Next
Convert2PDF.ConvertToPDF
End Sub

Sub RKJM(ByVal control As IRibbonControl)
On Error Resume Next
Download.DownRKJM
End Sub
Sub RKT(ByVal control As IRibbonControl)
On Error Resume Next
Download.DownRKT
End Sub

Sub Optionals(ByVal control As IRibbonControl)
On Error Resume Next
Download.DownOptionals
End Sub

' Untuk menentukan apakah tab tertentu aktif atau tidak
Sub GetEnabled(control As IRibbonControl, ByRef MakeVisible)
    Dim sheetName As String
    sheetName = "DEV"
    On Error Resume Next
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets(sheetName)
    Select Case control.Id
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
    Select Case control.Id
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
        Case "customGroup7":            MakeVisible = ws.Range("L34").value
        Case "customGroup8":            MakeVisible = ws.Range("L35").value
        Case "customGroup9":            MakeVisible = ws.Range("L36").value
        Case "customGroup10":           MakeVisible = ws.Range("L37").value
        Case "customGroup11":           MakeVisible = ws.Range("L38").value
        
        Case "Dash":                    MakeVisible = ws.Range("L39").value
        Case "Update":                  MakeVisible = ws.Range("L40").value
        Case "Upload":                  MakeVisible = ws.Range("L41").value
        Case "PetaBenahi":              MakeVisible = ws.Range("L42").value
        Case "LembarRKT":               MakeVisible = ws.Range("L43").value
        Case "LembarRKAS":              MakeVisible = ws.Range("L44").value
        Case "PrintView":               MakeVisible = ws.Range("L45").value
        Case "Saved":                   MakeVisible = ws.Range("L46").value
        
        Case "Data":                    MakeVisible = ws.Range("L47").value
        Case "DataRapat":               MakeVisible = ws.Range("L48").value
        Case "Matrix":                  MakeVisible = ws.Range("L49").value
        Case "HarsatBarjas":            MakeVisible = ws.Range("L50").value
        Case "HarsatModal":             MakeVisible = ws.Range("L51").value
        
        Case "AnalisisGugus":           MakeVisible = ws.Range("L52").value
        Case "AnalisisBuku":            MakeVisible = ws.Range("L53").value
        Case "AnalisisEkskul":          MakeVisible = ws.Range("L54").value
        Case "AnalisisHonor":           MakeVisible = ws.Range("L55").value
        
        Case "RKASROB":                 MakeVisible = ws.Range("L56").value
        Case "RKASPerTahap":            MakeVisible = ws.Range("L57").value
        Case "RKASSNP":                 MakeVisible = ws.Range("L58").value
        Case "RKASSIPD":                MakeVisible = ws.Range("L59").value
        Case "KomponenBOS":             MakeVisible = ws.Range("L60").value
        Case "RekonSaldo":              MakeVisible = ws.Range("L61").value
        
        Case "RBK":                     MakeVisible = ws.Range("L62").value
        Case "ReloadRBK":               MakeVisible = ws.Range("L63").value
        Case "Planning1":               MakeVisible = ws.Range("L64").value
        Case "Planning2":               MakeVisible = ws.Range("L65").value
        Case "PlanningTahun":           MakeVisible = ws.Range("L66").value
        
        Case "CoverRKAS":               MakeVisible = ws.Range("L67").value
        Case "CoverRKASPerubahan":      MakeVisible = ws.Range("L68").value
        Case "SKBendahara":             MakeVisible = ws.Range("L69").value
        Case "SKTimBOS":                MakeVisible = ws.Range("L70").value
        Case "SKTimPBJSekolah":         MakeVisible = ws.Range("L71").value
        Case "SKTAS":                   MakeVisible = ws.Range("L72").value
        Case "PenyusunanRKAS":          MakeVisible = ws.Range("L73").value
        Case "BelanjaModal":            MakeVisible = ws.Range("L74").value
        Case "PenggunaanDana":          MakeVisible = ws.Range("L75").value
        Case "LembarPengesahan":        MakeVisible = ws.Range("L76").value
        Case "Verval":                  MakeVisible = ws.Range("L77").value
        Case "RKJM":                    MakeVisible = ws.Range("L78").value
        Case "RKT":                     MakeVisible = ws.Range("L79").value
        Case "Optionals":               MakeVisible = ws.Range("L80").value
        Case Else:                      MakeVisible = False
    End Select
End Sub
' Subroutine untuk menentukan visibilitas tab berdasarkan nilai variabel
Sub GetLabel(control As IRibbonControl, ByRef MakeLabel)
    Dim sheetName As String
    sheetName = "DEV"
    On Error Resume Next
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets(sheetName)
    Select Case control.Id
        Case "Optionals":               MakeLabel = ws.Range("K7").value
        Case Else:                      MakeVisible = False
    End Select
End Sub

