Sub InsertRowsEveryNAndCopyValues(startRow As Long, interval As Long)
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim k As Long
    Dim maxK As Long
    Dim insertAt As Long
    Dim totalRows As Long
    Const SHEET_NAME As String = "Sheet1"

    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(SHEET_NAME)
    On Error GoTo 0
    
    If ws Is Nothing Then
        MsgBox "Worksheet '" & SHEET_NAME & "' not found.", vbExclamation
        Exit Sub
    End If

    If ws.ProtectContents Or ws.ProtectDrawingObjects Or ws.ProtectScenarios Then
        MsgBox "Worksheet '" & SHEET_NAME & "' is protected. Unprotect it first.", vbExclamation
        Exit Sub
    End If

    totalRows = ws.Rows.Count
    lastRow = ws.Cells(totalRows, "A").End(xlUp).Row

    If startRow > lastRow Then
        MsgBox "startRow (" & startRow & ") is below any used data (lastRow = " & lastRow & ").", vbInformation
        Exit Sub
    End If

    maxK = Int((lastRow - startRow) / interval)

    For k = maxK To 0 Step -1
        insertAt = startRow + (k * interval)
        
        If insertAt >= 1 And insertAt < totalRows Then
            On Error Resume Next
            ws.Rows(insertAt).Insert Shift:=xlDown
            If Err.Number = 0 Then
                ' Copy values from the row below into the inserted row (columns A to E)
                ws.Range("A" & insertAt + 1 & ":E" & insertAt + 1).Copy
                ws.Range("A" & insertAt).PasteSpecial Paste:=xlPasteValues
            Else
                MsgBox "Error inserting at row " & insertAt & ": " & Err.Description, vbExclamation
                Err.Clear
            End If
            On Error GoTo 0
        End If
    Next k

    Application.CutCopyMode = False
End Sub
