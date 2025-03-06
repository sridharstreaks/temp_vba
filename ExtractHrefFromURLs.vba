Sub ExtractHrefFromURLs()
    Dim wsName As String, colName As String
    Dim startRow As Long, endRow As Long, lastRow As Long, i As Long
    Dim urlToProcess As String, extractedHref As String
    Dim IE As Object, elem As Object
    Dim endRowInput As String
    Dim ws As Worksheet
    Dim startTime As Single

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
    IE.Visible = False   ' Change to True if you want to see the navigation
    
    ' Loop through each URL in the specified range
    For i = startRow To endRow
        urlToProcess = ws.Cells(i, ws.Range(colName & "1").Column).Value
        If urlToProcess <> "" Then
            IE.Navigate urlToProcess
            
            ' Wait for IE to finish loading
            Do While IE.Busy Or IE.ReadyState <> 4
                DoEvents
            Loop
            
            ' Additionally, wait for the document readyState to be "complete"
            On Error Resume Next
            Do While IE.Document.readyState <> "complete"
                DoEvents
            Loop
            On Error GoTo 0
            
            ' Wait for the specific element to appear (timeout after 10 seconds)
            startTime = Timer
            Set elem = Nothing
            Do
                On Error Resume Next
                Set elem = IE.Document.querySelector("h1.title a")
                On Error GoTo 0
                If Not elem Is Nothing Then Exit Do
                If Timer - startTime > 10 Then Exit Do ' Timeout after 10 seconds
                DoEvents
            Loop
            
            If Not elem Is Nothing Then
                extractedHref = elem.href
            Else
                extractedHref = "Not found"
            End If
            
            ' Write the extracted link in the cell immediately to the right of the URL cell
            ws.Cells(i, ws.Range(colName & "1").Column + 1).Value = extractedHref
        End If
    Next i
    
    ' Clean up
    IE.Quit
    Set IE = Nothing
    
    MsgBox "Extraction complete.", vbInformation
End Sub
