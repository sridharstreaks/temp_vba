Sub DeleteInsertedRows()
    Dim ws      As Worksheet
    Dim lastRow As Long
    Dim r       As Long
    
    ' Change "Sheet1" to your actual sheet name:
    Set ws = ThisWorkbook.Worksheets("Sheet1")
    
    ' Find the last used row in column A (change "A" if you want a different anchor column)
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    
    ' Loop from the second‐to‐last row up to row 1:
    For r = lastRow - 1 To 1 Step -1
        ' Compare columns A:E of row r vs. row (r+1)
        If ws.Cells(r, "A").Value = ws.Cells(r + 1, "A").Value _
           And ws.Cells(r, "B").Value = ws.Cells(r + 1, "B").Value _
           And ws.Cells(r, "C").Value = ws.Cells(r + 1, "C").Value _
           And ws.Cells(r, "D").Value = ws.Cells(r + 1, "D").Value _
           And ws.Cells(r, "E").Value = ws.Cells(r + 1, "E").Value Then
               
            ws.Rows(r).Delete Shift:=xlUp
        End If
    Next r
End Sub
