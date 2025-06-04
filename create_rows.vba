Sub PasteXlookupFormulasInInsertedRows()
    Dim ws         As Worksheet
    Dim lookupWS   As Worksheet
    Dim lastRow    As Long
    Dim lastCol    As Long
    Dim r          As Long
    Dim c          As Long
    Dim matchAE    As Boolean
    
    '──— Change these sheet names as needed ──—
    Set ws = ThisWorkbook.Worksheets("Sheet1")   ' sheet where rows were inserted
    Set lookupWS = ThisWorkbook.Worksheets("Sheet2")  ' sheet that holds columns A, C, E for XLOOKUP
    
    ' Find the last used row in column A of ws:
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    ' Find the last filled column in row 1 of ws:
    lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
    
    ' Loop from row 1 down to the second-to-last row:
    ' If in ws, A:E of row r = A:E of row r+1, then r is one of our “inserted” rows.
    For r = 1 To lastRow - 1
        matchAE = _
            (ws.Cells(r, "A").Value = ws.Cells(r + 1, "A").Value) And _
            (ws.Cells(r, "B").Value = ws.Cells(r + 1, "B").Value) And _
            (ws.Cells(r, "C").Value = ws.Cells(r + 1, "C").Value) And _
            (ws.Cells(r, "D").Value = ws.Cells(r + 1, "D").Value) And _
            (ws.Cells(r, "E").Value = ws.Cells(r + 1, "E").Value)
        
        If matchAE Then
            ' r is an inserted row.  Paste the XLOOKUP formula from col F to lastCol:
            For c = 6 To lastCol
                ' Using R1C1 style so that:
                '   • RC3   → column C of THIS (inserted) row (e.g. C2 if r=2)
                '   • R1C   → row 1 of THIS column (e.g. F1, G1, etc.)
                ' Whole‐column references to lookupWS:
                '   `'Sheet2'!C:C`, `'Sheet2'!A:A`, `'Sheet2'!E:E`
                ws.Cells(r, c).FormulaR1C1 = _
                    "=XLOOKUP(1," & _
                        "('" & lookupWS.Name & "'!C:C=RC3)*(" & _
                        "'" & lookupWS.Name & "'!A:A=R1C)" & _
                    ",'" & lookupWS.Name & "'!E:E)"
            Next c
        End If
    Next r
End Sub
