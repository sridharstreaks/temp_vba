Sub InsertRowsWithValuesAndFormulas(startRow As Long, interval As Long)
    Dim ws         As Worksheet
    Dim wsLookup   As Worksheet
    Dim lastRow    As Long
    Dim lastCol    As Long
    Dim totalRows  As Long
    Dim maxK       As Long
    Dim k          As Long
    Dim insertAt   As Long
    Dim headerCell As Range
    Dim formulaStr As String
    
    Const SHEET_NAME As String = "Sheet1"
    
    '――――――――――――――――――――――――――――――――――――――――――――――――――――――――――
    ' 1) Get “current” worksheet (where we insert)
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(SHEET_NAME)
    On Error GoTo 0
    
    If ws Is Nothing Then
        MsgBox "Worksheet '" & SHEET_NAME & "' not found.", vbCritical
        Exit Sub
    End If
    
    ' 2) Ensure the sheet isn’t protected
    If ws.ProtectContents Or ws.ProtectDrawingObjects Or ws.ProtectScenarios Then
        MsgBox "Worksheet '" & SHEET_NAME & "' is protected. Unprotect it first.", vbExclamation
        Exit Sub
    End If
    
    ' 3) Determine “next” sheet for lookup (must exist)
    If ws.Index = ThisWorkbook.Worksheets.Count Then
        MsgBox "No worksheet after '" & SHEET_NAME & "'. Cannot find lookup sheet.", vbExclamation
        Exit Sub
    End If
    Set wsLookup = ThisWorkbook.Worksheets(ws.Index + 1)
    
    ' 4) Figure out where current data ends in column A (anchor column for “last used row”)
    totalRows = ws.Rows.Count
    lastRow   = ws.Cells(totalRows, "A").End(xlUp).Row
    
    If startRow > lastRow Then
        MsgBox "startRow (" & startRow & ") is below any used data (lastRow = " & lastRow & ").", vbInformation
        Exit Sub
    End If
    
    ' 5) Find last header column in row 1 (we assume headers run from I1 → some lastCol without gaps)
    lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
    If lastCol < 9 Then
        MsgBox "Row 1 does not have any headers from column I onward.", vbExclamation
        Exit Sub
    End If
    
    ' 6) Compute how many insertions “fit” between startRow and lastRow
    maxK = Int((lastRow - startRow) / interval)
    
    ' 7) Loop backwards so that inserting rows doesn't shift future insert positions
    For k = maxK To 0 Step -1
        insertAt = startRow + (k * interval)
        
        ' Make sure insertAt is a valid row number
        If insertAt >= 1 And insertAt < totalRows Then
            On Error Resume Next
            ws.Rows(insertAt).Insert Shift:=xlDown
            If Err.Number <> 0 Then
                MsgBox "Could not insert row at " & insertAt & ": " & Err.Description, vbExclamation
                Err.Clear
            Else
                ' ―――――――――――――――――――――――――――――――――――――――――――
                ' After inserting the blank row at 'insertAt', do:
                '   • Copy A:D from (insertAt+1) → (insertAt)
                '   • Put "blah" in column E of the inserted row
                '   • Leave F, G, H empty
                '   • Paste XLOOKUP formula from I → lastCol
                With ws
                    ' Copy A:D from the row _below_ (row = insertAt+1) into inserted row
                    .Range("A" & insertAt + 1 & ":D" & insertAt + 1).Copy
                    .Range("A" & insertAt & ":D" & insertAt).PasteSpecial Paste:=xlPasteValues
                    Application.CutCopyMode = False
                    
                    ' Set column E to the literal "blah"
                    .Cells(insertAt, "E").Value = "blah"
                    
                    ' F, G, H: do nothing (they remain blank)
                    
                    ' Build & paste the XLOOKUP formula from I through lastCol:
                    '   =XLOOKUP(
                    '       1,
                    '       (wsLookup!C:C = C<row>) * (wsLookup!A:A = <HeaderCell>),
                    '       wsLookup!E:E,
                    '       "",
                    '       0
                    '    )
                    '
                    ' – C<row> is .Cells(insertAt, "C")
                    ' – <HeaderCell> is .Cells(1, col) (e.g. I1, J1, K1, …)
                    ' – wsLookup.Name is used in the sheet reference
                    '
                    Dim lookupName As String
                    lookupName = wsLookup.Name
                    If InStr(1, lookupName, " ") > 0 Then
                        ' wrap in single quotes if sheet name has spaces
                        lookupName = "'" & lookupName & "'"
                    End If
                    
                    Dim col As Long
                    For col = 9 To lastCol
                        ' reference to C<insertAt> (no $ so that column stays as "C", row is the inserted row)
                        ' reference to header in row 1: e.g. I$1, J$1, K$1 (column is relative, row absolute)
                        Dim thisHeaderRef As String
                        thisHeaderRef = .Cells(1, col).Address(False, True)   ' e.g. "I$1", "J$1", etc.
                        
                        formulaStr = "=XLOOKUP( " & _
                                     "1, " & _
                                     "(" & lookupName & "!C:C = " & .Cells(insertAt, "C").Address(False, False) & ") * " & _
                                     "(" & lookupName & "!A:A = " & thisHeaderRef & "), " & _
                                     lookupName & "!E:E, " & _
                                     """" & """" & ", 0" & _
                                     ")"
                                     
                        .Cells(insertAt, col).Formula = formulaStr
                    Next col
                End With
            End If
            On Error GoTo 0
        End If
    Next k
End Sub
