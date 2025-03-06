Sub ExtractHrefFromURLs()
    Dim wsName As String, colName As String
    Dim startRow As Long, endRow As Long, lastRow As Long, i As Long
    Dim urlToProcess As String, extractedHref As String
    Dim IE As Object, elem As Object
    Dim endRowInput As String
    Dim ws As Worksheet
    
    ' Prompt for inputs
    wsName = InputBox("Enter the worksheet name containing the URLs:")
    colName = InputBox("Enter the column letter containing the URLs (e.g., A):")
    startRow = CLng(InputBox("Enter the starting row number:"))
    endRowInput = InputBox("Enter the ending row number (leave blank for last filled row):")
    
    Set ws = ThisWorkbook.Worksheets(wsName)
    
    ' Determine ending row if left blank
    If Trim(endRowInput) = "" Then
        lastRow = ws.Cells(ws.Rows.Count, colName).End(xlUp).Row
        endRow = lastRow
    Else
        endRow = CLng(endRowInput)
    End If
    
    ' Create and configure Internet Explorer instance
    Set IE = CreateObject("InternetExplorer.Application")
    IE.Visible = False   ' Change to True if you wish to see the navigation
    
    ' Loop through each URL in the specified range
    For i = startRow To endRow
        urlToProcess = ws.Cells(i, ws.Range(colName & "1").Column).Value
        If urlToProcess <> "" Then
            IE.Navigate urlToProcess
            ' Wait for page to load completely
            Do While IE.Busy Or IE.ReadyState <> 4
                DoEvents
            Loop
            
            On Error Resume Next
            ' Use querySelector to get the anchor inside h1.title
            Set elem = IE.Document.querySelector("h1.title a")
            If Not elem Is Nothing Then
                extractedHref = elem.href
            Else
                extractedHref = "Not found"
            End If
            On Error GoTo 0
            
            ' Write the extracted link in the cell immediately to the right of the URL cell
            ws.Cells(i, ws.Range(colName & "1").Column + 1).Value = extractedHref
        End If
    Next i
    
    ' Clean up
    IE.Quit
    Set IE = Nothing
    
    MsgBox "Extraction complete.", vbInformation
End Sub
