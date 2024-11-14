Sub SaveChartAudit()
    Dim wsAudit As Worksheet, wsCharts As Worksheet
    Dim lastRowAudit As Long, lastRowCharts As Long
    Dim i As Long
    
    Set wsAudit = ThisWorkbook.Sheets("Chart Audit")
    Set wsCharts = ThisWorkbook.Sheets("Audited Charts")
    
    lastRowAudit = wsAudit.Cells(wsAudit.Rows.Count, 1).End(xlUp).Row
    MsgBox "Last row in Chart Audit: " & lastRowAudit ' Check if last row is correct
    
    For i = 2 To lastRowAudit
        If wsAudit.Cells(i, 1).Value <> "" Then
            lastRowCharts = wsCharts.Cells(wsCharts.Rows.Count, 1).End(xlUp).Row + 1
            MsgBox "Copying data from row: " & i ' Check which row is being copied
            
            wsCharts.Cells(lastRowCharts, 1).Value = wsAudit.Cells(i, 1).Value
            wsCharts.Cells(lastRowCharts, 2).Value = wsAudit.Cells(i, 2).Value
            wsCharts.Cells(lastRowCharts, 3).Value = wsAudit.Cells(i, 3).Value
            wsCharts.Cells(lastRowCharts, 4).Value = wsAudit.Cells(i, 4).Value
            wsCharts.Cells(lastRowCharts, 5).Value = wsAudit.Cells(i, 5).Value
        End If
    Next i
    
    MsgBox "Data saved successfully for all rows!"
End Sub
