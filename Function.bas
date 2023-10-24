Sub BuatSheetBaru()
    ' Cara Penggunaan: Jalankan makro ini untuk membuat sheet baru
    Sheets.Add
End Sub

Sub HapusSheet(NamaSheet As String)
    ' Cara Penggunaan: Jalankan makro ini dengan menyebutkan nama sheet yang ingin dihapus
    On Error Resume Next
    Sheets(NamaSheet).Delete
    On Error GoTo 0
End Sub

Sub GantiNamaSheet(NamaLama As String, NamaBaru As String)
    ' Cara Penggunaan: Jalankan makro ini dengan menyebutkan nama sheet yang ingin diganti dan nama baru
    On Error Resume Next
    Sheets(NamaLama).Name = NamaBaru
    On Error GoTo 0
End Sub

Sub PindahSheet(NamaSheet As String, IndexTujuan As Integer)
    ' Cara Penggunaan: Jalankan makro ini dengan menyebutkan nama sheet dan indeks tempat tujuan
    On Error Resume Next
    Sheets(NamaSheet).Move Before:=Sheets(IndexTujuan)
    On Error GoTo 0
End Sub

Sub SalinSheet(NamaSheet As String)
    ' Cara Penggunaan: Jalankan makro ini dengan menyebutkan nama sheet yang ingin disalin
    Sheets(NamaSheet).Copy
End Sub

Sub LindungiSheet(NamaSheet As String, KataSandi As String)
    ' Cara Penggunaan: Jalankan makro ini dengan menyebutkan nama sheet dan kata sandi
    Sheets(NamaSheet).Protect Password:=KataSandi
End Sub

Sub HapusPerlindunganSheet(NamaSheet As String, KataSandi As String)
    ' Cara Penggunaan: Jalankan makro ini dengan menyebutkan nama sheet dan kata sandi
    Sheets(NamaSheet).Unprotect Password:=KataSandi
End Sub

Sub SembunyikanSheet(NamaSheet As String)
    ' Cara Penggunaan: Jalankan makro ini dengan menyebutkan nama sheet yang ingin disembunyikan
    Sheets(NamaSheet).Visible = xlSheetHidden
End Sub

Sub TampilkanSheet(NamaSheet As String)
    ' Cara Penggunaan: Jalankan makro ini dengan menyebutkan nama sheet yang ingin ditampilkan
    Sheets(NamaSheet).Visible = xlSheetVisible
End Sub

Sub PilihSemuaSheet()
    ' Cara Penggunaan: Jalankan makro ini untuk memilih semua sheet dalam workbook
    Sheets.Select
End Sub

Sub GantiWarnaTab(NamaSheet As String, Warna As Long)
    ' Cara Penggunaan: Jalankan makro ini dengan menyebutkan nama sheet dan warna yang diinginkan (misalnya RGB(255, 0, 0) untuk merah)
    Sheets(NamaSheet).Tab.Color = Warna
End Sub
