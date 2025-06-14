Sub MailMerged()
    Dim fso As Object, http As Object, stream As Object
    Dim wordApp As Object, wordDoc As Object
    Dim ws As Worksheet
    Dim basePath As String, tempFolder As String, tempFolderClean As String
    Dim outputFolder As String, filePath As String, pdfPath As String
    Dim downloadURL As String
    Dim header As String, value As String
    Dim i As Long, j As Long, lastCol As Long, lastRow As Long

    Set fso = CreateObject("Scripting.FileSystemObject")
    Set ws = ThisWorkbook.Sheets("mail")
    basePath = ThisWorkbook.Path & "\"
    tempFolder = basePath & "RBKdownload\"
    tempFolderClean = basePath & "RBKdownload"
    If Not fso.FolderExists(tempFolder) Then fso.CreateFolder tempFolder

    outputFolder = basePath & "GENERATE RBK 2025\"
    If Not fso.FolderExists(outputFolder) Then fso.CreateFolder outputFolder

    downloadURL = "https://docs.google.com/document/d/1O6gAYOr3B4CNybingzGrXNAPcsFR2MWUBAwRBR9T-nE/export?format=doc"
    filePath = tempFolder & "template.docx"

    Set http = CreateObject("WinHttp.WinHttpRequest.5.1")
    http.Open "GET", downloadURL, False
    http.Send

    If http.Status = 200 Then
        Set stream = CreateObject("ADODB.Stream")
        With stream
            .Type = 1
            .Open
            .Write http.responseBody
            .SaveToFile filePath, 2
            .Close
        End With
    Else
        MsgBox "Gagal mendownload file", vbCritical
        Exit Sub
    End If

    Set wordApp = CreateObject("Word.Application")
    wordApp.Visible = False

    lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row

    ' Loop setiap baris data mulai dari baris ke-2
    For i = 2 To lastRow
        Set wordDoc = wordApp.Documents.Open(filePath)
        
        For j = 1 To lastCol
            header = Trim(ws.Cells(1, j).value)
            value = ws.Cells(i, j).value

            With wordDoc.Content.Find
                .Text = "<<" & header & ">>"
                .Replacement.Text = value
                .Forward = True
                .Wrap = 1
                .Execute Replace:=2
            End With
        Next j

        pdfPath = outputFolder & "RBK_" & Format(Now, "yyyymmdd_hhnnss") & "_" & i & ".pdf"
        wordDoc.ExportAsFixedFormat OutputFileName:=pdfPath, ExportFormat:=17
        wordDoc.Close False
    Next i

    wordApp.Quit
    Set wordApp = Nothing

    On Error Resume Next
    If fso.FolderExists(tempFolderClean) Then
        fso.DeleteFolder tempFolderClean, True
    End If

    MsgBox "PDF berhasil dibuat sebanyak " & (lastRow - 1) & " file.", vbInformation
End Sub
