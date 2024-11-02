Sub LoadJsonData()
    Dim wb As Workbook
    Dim LoadDataSheet As Worksheet
    Dim QueryName As String
    Dim url As String
    
    On Error Resume Next

    ' Mengatur workbook dan sheet
    Set wb = ThisWorkbook
    Set LoadDataSheet = wb.Sheets("Sheet1")

    ' Konfigurasi yang dapat diubah
    QueryName = "MyJsonQuery"
    url = "https://white-bar-3fc0.arvib.workers.dev/2PACX-1vTXtiHnRqxjLfrwD3apjwv3NZ4HAlWVdEwxkArptz8CiBKdumW3STXKNdMLligsyBHxNhLBGTkdaoWz/12345"
    
    ' Menghapus data di Sheet1
    LoadDataSheet.Cells.Clear

    ' Menambahkan kueri baru
    wb.Queries.Add Name:=QueryName, Formula:= _
        "let" & vbCrLf & _
        "    Source = Json.Document(Web.Contents(""" & url & """))," & vbCrLf & _
        "    RecordAsTable = Record.ToTable(Source{0})," & vbCrLf & _
        "    PromotedHeaders = Table.PromoteHeaders(Table.Transpose(RecordAsTable), [PromoteAllScalars=true])" & vbCrLf & _
        "in" & vbCrLf & _
        "    PromotedHeaders"

    ' Memuat data ke Sheet1 di sel A1
    LoadQuery QueryName, LoadDataSheet

    ' Menghapus kueri setelah pemuatan
    wb.Queries(QueryName).Delete
    Dim conn As WorkbookConnection
    For Each conn In wb.Connections
        conn.Delete
    Next conn

End Sub

Private Sub LoadQuery(ByVal QueryName As String, ByVal LoadDataSheet As Worksheet)
    With LoadDataSheet.ListObjects.Add(SourceType:=0, source:= _
        "OLEDB;Provider=Microsoft.Mashup.OleDb.1;Data Source=$Workbook$;Location=" & QueryName & ";Extended Properties=""""" _
        , Destination:=LoadDataSheet.Range("$A$1")).queryTable
        .CommandType = xlCmdSql
        .CommandText = Array("SELECT * FROM [" & QueryName & "]")
        .BackgroundQuery = False
        .AdjustColumnWidth = True
        .Refresh BackgroundQuery:=False
        .Delete
    End With
End Sub
