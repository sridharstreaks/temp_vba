Sub InsertRowsEveryN(startRow As Long, interval As Long)
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim k As Long
    Dim maxK As Long
    Dim insertAt As Long
    
    ' Change "Sheet1" to whichever sheet you want this to run on:
    Set ws = ThisWorkbook.Worksheets("Sheet1")
    
    ' Find the last used row in column A (you can switch to a different column if needed)
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    
    ' If startRow is already beyond the lastRow, nothing to do:
    If startRow > lastRow Then Exit Sub
    
    ' Compute how many insertions of size "interval" fit between startRow and lastRow:
    '   For k = 0 To maxK, rowToInsert = startRow + k * interval
    maxK = Int((lastRow - startRow) / interval)
    
    ' Loop backwards so that inserting rows doesn’t affect positions of the ones not yet inserted:
    For k = maxK To 0 Step -1
        insertAt = startRow + (k * interval)
        ws.Rows(insertAt).Insert Shift:=xlDown
    Next k
End Sub
