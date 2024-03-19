Sub SendDataToGoogleSheet()
    Dim URL As String
    Dim HTTPReq As Object
    Dim JSONString As String
    Dim RangeData As Range
    Dim DataArray As Variant
    Dim i As Long

    ' URL for Google Sheets REST API
    URL = "https://script.google.com/macros/s/AKfycby9dV6iBiXT8r5wLFiUQYSrFVG5Q3V0AcEAL9E61vEA9YZzysdQRtUfpXRlh8Z5g6K-/exec?sheet=Sheet2&range=A:H"

    ' Set the data range from Excel
    Set RangeData = ThisWorkbook.Sheets("DEV").Range("AA2:AH" & ThisWorkbook.Sheets("DEV").Cells(ThisWorkbook.Sheets("DEV").Rows.Count, "AA").End(xlUp).Row)

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
    HTTPReq.send JSONString

    ' Show the result message
    MsgBox "Data has been successfully sent to Google Sheets.", vbInformation
End Sub
