Sub SaveChartAudit()
    Dim wsAudit As Worksheet, wsCharts As Worksheet
    Dim lastRowAudit As Long, lastRowCharts As Long
    Dim i As Long

    Set wsAudit = ThisWorkbook.Sheets("Chart Audit")
    Set wsCharts = ThisWorkbook.Sheets("Audited Charts")
    
    lastRowAudit = wsAudit.Cells(wsAudit.Rows.Count, 1).End(xlUp).Row
    lastRowCharts = wsCharts.Cells(wsCharts.Rows.Count, 1).End(xlUp).Row + 1
    
    For i = 2 To lastRowAudit
        If wsAudit.Cells(i, 1).Value <> "" Then ' Check if there is data in the row
            ' Copy data from Chart Audit to Audited Charts
            wsCharts.Cells(lastRowCharts, 1).Value = wsAudit.Cells(i, 1).Value ' Patient Name
            wsCharts.Cells(lastRowCharts, 2).Value = wsAudit.Cells(i, 2).Value ' Document
            wsCharts.Cells(lastRowCharts, 3).Value = wsAudit.Cells(i, 3).Value ' Complete Information
            wsCharts.Cells(lastRowCharts, 4).Value = wsAudit.Cells(i, 4).Value ' Correct Information
            wsCharts.Cells(lastRowCharts, 5).Value = wsAudit.Cells(i, 5).Value ' Remarks
            
            lastRowCharts = lastRowCharts + 1 ' Move to the next row in Audited Charts
        End If
    Next i
    
    MsgBox "Data saved successfully!"
End Sub
