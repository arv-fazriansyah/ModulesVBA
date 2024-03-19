Sub SendDataToGoogleSheet()
    Dim URL As String
    Dim HTTPReq As Object
    Dim JSONString As String
    Dim RangeData As Range
    Dim DataArray As Variant
    Dim i As Long
    Dim SheetName As String
    Dim StartColumn As String
    Dim EndColumn As String

    ' URL for Google Sheets REST API
    URL = "https://" & SubPath & "." & Author & ".eu.org/" & "send"

    ' Set the sheet name and data range
    SheetName = "DEV"
    StartColumn = "AA"
    EndColumn = "AH"

    ' Set the data range from Excel
    With ThisWorkbook.Sheets(SheetName)
        Set RangeData = .Range(StartColumn & "2:" & EndColumn & .Cells(.Rows.Count, StartColumn).End(xlUp).Row)
    End With

    ' Convert the data range to an array
    DataArray = RangeData.Value

    ' Create the JSON string
    JSONString = "{""values"": ["
    For i = 1 To UBound(DataArray)
        JSONString = JSONString & "["
        For j = 1 To UBound(DataArray, 2)
            JSONString = JSONString & """" & DataArray(i, j) & """"
            If j <> UBound(DataArray, 2) Then
                JSONString = JSONString & ","
            End If
        Next j
        JSONString = JSONString & "]"
        If i <> UBound(DataArray) Then
            JSONString = JSONString & ","
        End If
    Next i
    JSONString = JSONString & "]}"

    ' Create the WinHttpRequest object
    Set HTTPReq = CreateObject("WinHttp.WinHttpRequest.5.1")

    ' Send the POST request
    HTTPReq.Open "POST", URL, False
    HTTPReq.setRequestHeader "Content-Type", "application/json"
    HTTPReq.Send JSONString

    ' Show the result message
    MsgBox "Data has been successfully sent to Google Sheets.", vbInformation
End Sub

