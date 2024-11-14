Sub SaveChartAudit()
       1 Dim wsAudit As Worksheet, wsCharts As Worksheet
       Set wsAudit = ThisWorkbook.Sheets("Chart Audit")
       Set wsCharts = ThisWorkbook.Sheets("Audited Charts")
       Dim lastRow As Long
       lastRow = wsCharts.Cells(wsCharts.Rows.Count, 1).End(xlUp).Row + 1

       ' Copy data from Chart Audit to Audited Charts
       wsCharts.Cells(lastRow, 1).Value = wsAudit.Cells(2, 1).Value ' Patient Name
       wsCharts.Cells(lastRow, 2).Value = wsAudit.Cells(2, 2).Value ' Document
       wsCharts.Cells(lastRow, 3).Value = wsAudit.Cells(2, 3).Value ' Complete Information
       wsCharts.Cells(lastRow, 4).Value = wsAudit.Cells(2, 4).Value ' Correct Information
       wsCharts.Cells(lastRow, 5).Value = wsAudit.Cells(2, 5).Value ' Remarks

       MsgBox "Data saved successfully!" 1
   End Sub
