Sub UpdateDashboard()
    Dim wsCharts As Worksheet, wsDashboard As Worksheet
    Dim lastRowCharts As Long, lastRowDashboard As Long
    Dim i As Long
    
    Set wsCharts = ThisWorkbook.Sheets("Audited Charts")
    Set wsDashboard = ThisWorkbook.Sheets("Dashboard")
    
    lastRowCharts = wsCharts.Cells(wsCharts.Rows.Count, 1).End(xlUp).Row
    lastRowDashboard = wsDashboard.Cells(wsDashboard.Rows.Count, 1).End(xlUp).Row + 1
    
    ' Clear existing summary in Dashboard, but not if there's missing patient name
    If wsDashboard.Cells(2, 1).Value <> "" Then
        wsDashboard.Range("A2:H" & lastRowDashboard).ClearContents
    End If
    
    For i = 2 To lastRowCharts
        If wsCharts.Cells(i, 1).Value <> "" Then
            ' Update dashboard with the data from Audited Charts
            wsDashboard.Cells(lastRowDashboard, 1).Value = wsCharts.Cells(i, 2).Value ' Document
            wsDashboard.Cells(lastRowDashboard, 2).Value = Application.CountIf(wsCharts.Range("B2:B" & lastRowCharts), wsCharts.Cells(i, 2).Value) ' Count
            wsDashboard.Cells(lastRowDashboard, 3).Value = Application.CountIf(wsCharts.Range("C2:C" & lastRowCharts), "Yes") ' Complete
            wsDashboard.Cells(lastRowDashboard, 4).Value = Application.CountIf(wsCharts.Range("C2:C" & lastRowCharts), "No") ' Incomplete
            wsDashboard.Cells(lastRowDashboard, 5).Value = Application.CountIf(wsCharts.Range("D2:D" & lastRowCharts), "Yes") ' Correct
            wsDashboard.Cells(lastRowDashboard, 6).Value = Application.CountIf(wsCharts.Range("D2:D" & lastRowCharts), "No") ' Incorrect
            wsDashboard.Cells(lastRowDashboard, 7).Formula = "=IF(B" & lastRowDashboard & ">0, D" & lastRowDashboard & "/B" & lastRowDashboard & ", 0)" ' % Incomplete
            wsDashboard.Cells(lastRowDashboard, 8).Formula = "=IF(B" & lastRowDashboard & ">0, F" & lastRowDashboard & "/B" & lastRowDashboard & ", 0)" ' % Incorrect
            
            lastRowDashboard = lastRowDashboard + 1
        End If
    Next i

    MsgBox "Dashboard updated successfully!"
End Sub
