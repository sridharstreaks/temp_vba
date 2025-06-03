Sub InsertRowsEveryN(startRow As Long, interval As Long)
    Dim ws       As Worksheet
    Dim lastRow  As Long
    Dim k        As Long
    Dim maxK     As Long
    Dim insertAt As Long
    Dim totalRows As Long
    
    '──— Change this to your actual sheet name ──—
    Const SHEET_NAME As String = "Sheet1"
    
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(SHEET_NAME)
    On Error GoTo 0
    
    If ws Is Nothing Then
        MsgBox "Worksheet '" & SHEET_NAME & "' not found.", vbExclamation
        Exit Sub
    End If
    
    ' If the sheet is protected, Insert will fail—unprotect or bail out:
    If ws.ProtectContents Or ws.ProtectDrawingObjects Or ws.ProtectScenarios Then
        MsgBox "Worksheet '" & SHEET_NAME & "' is protected. Unprotect it first.", vbExclamation
        Exit Sub
    End If
    
    totalRows = ws.Rows.Count   ' typically 1,048,576 on modern Excel
    
    ' Find the last used row in column A (change "A" if you need a different column)
    lastRow = ws.Cells(totalRows, "A").End(xlUp).Row
    
    ' If startRow is beyond lastRow, nothing to do
    If startRow > lastRow Then
        MsgBox "startRow (" & startRow & ") is below any used data (lastRow = " & lastRow & ").", vbInformation
        Exit Sub
    End If
    
    ' Calculate how many insertions fit between startRow and lastRow
    maxK = Int((lastRow - startRow) / interval)
    
    ' Loop backwards: k = maxK down to 0
    For k = maxK To 0 Step -1
        insertAt = startRow + (k * interval)
        
        ' Make sure insertAt is between 1 and totalRows
        If insertAt >= 1 And insertAt <= totalRows Then
            On Error Resume Next
            ws.Rows(insertAt).Insert Shift:=xlDown
            If Err.Number <> 0 Then
                ' If even a valid‐looking row causes an error, show it
                MsgBox "Error inserting at row " & insertAt & ": " & Err.Description, vbExclamation
                Err.Clear
            End If
            On Error GoTo 0
        Else
            ' If insertAt is out of bounds, skip it
        End If
    Next k
End Sub
