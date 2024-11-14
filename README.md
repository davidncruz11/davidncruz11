Sub UpdateDashboard()
    Dim wsAudit As Worksheet, wsDashboard As Worksheet
    Dim lastRow As Long, rowCount As Long
    Dim docName As String, countDoc As Long, countComplete As Long, countIncomplete As Long, countCorrect As Long, countIncorrect As Long
    Dim i As Long
    
    ' Set references to the relevant sheets
    Set wsAudit = ThisWorkbook.Sheets("Audited Charts")
    Set wsDashboard = ThisWorkbook.Sheets("Dashboard")
    
    ' Find the last row with data in the Audited Charts sheet
    lastRow = wsAudit.Cells(wsAudit.Rows.Count, 1).End(xlUp).Row
    
    ' Clear previous data from the Dashboard (optional)
    wsDashboard.Range("A2:H" & wsDashboard.Cells(wsDashboard.Rows.Count, 1).End(xlUp).Row).ClearContents
    
    ' Loop through each row in the Audited Charts sheet
    For i = 2 To lastRow
        docName = wsAudit.Cells(i, 2).Value ' Get the document type from Column B (Document)
        
        ' Check if the document is already in the Dashboard
        rowCount = Application.WorksheetFunction.CountIf(wsDashboard.Range("A:A"), docName)
        
        ' If the document is already in the Dashboard, update its counts
        If rowCount > 0 Then
            countDoc = wsDashboard.Cells(Application.WorksheetFunction.Match(docName, wsDashboard.Range("A:A"), 0), 2).Value
            countComplete = wsDashboard.Cells(Application.WorksheetFunction.Match(docName, wsDashboard.Range("A:A"), 0), 3).Value
            countIncomplete = wsDashboard.Cells(Application.WorksheetFunction.Match(docName, wsDashboard.Range("A:A"), 0), 4).Value
            countCorrect = wsDashboard.Cells(Application.WorksheetFunction.Match(docName, wsDashboard.Range("A:A"), 0), 5).Value
            countIncorrect = wsDashboard.Cells(Application.WorksheetFunction.Match(docName, wsDashboard.Range("A:A"), 0), 6).Value
        Else
            ' If the document isn't in the Dashboard yet, initialize counts
            countDoc = 0
            countComplete = 0
            countIncomplete = 0
            countCorrect = 0
            countIncorrect = 0
        End If
        
        ' Update the counts based on the data in Audited Charts
        If wsAudit.Cells(i, 3).Value = "Yes" Then
            countComplete = countComplete + 1
        Else
            countIncomplete = countIncomplete + 1
        End If
        
        If wsAudit.Cells(i, 4).Value = "Yes" Then
            countCorrect = countCorrect + 1
        Else
            countIncorrect = countIncorrect + 1
        End If
        
        ' Update the Dashboard with the new counts
        If rowCount > 0 Then
            wsDashboard.Cells(Application.WorksheetFunction.Match(docName, wsDashboard.Range("A:A"), 0), 2).Value = countDoc + 1
            wsDashboard.Cells(Application.WorksheetFunction.Match(docName, wsDashboard.Range("A:A"), 0), 3).Value = countComplete
            wsDashboard.Cells(Application.WorksheetFunction.Match(docName, wsDashboard.Range("A:A"), 0), 4).Value = countIncomplete
            wsDashboard.Cells(Application.WorksheetFunction.Match(docName, wsDashboard.Range("A:A"), 0), 5).Value = countCorrect
            wsDashboard.Cells(Application.WorksheetFunction.Match(docName, wsDashboard.Range("A:A"), 0), 6).Value = countIncorrect
        Else
            ' If the document is not already in the Dashboard, add it
            wsDashboard.Cells(wsDashboard.Cells(wsDashboard.Rows.Count, 1).End(xlUp).Row + 1, 1).Value = docName
            wsDashboard.Cells(wsDashboard.Cells(wsDashboard.Rows.Count, 1).End(xlUp).Row, 2).Value = countDoc + 1
            wsDashboard.Cells(wsDashboard.Cells(wsDashboard.Rows.Count, 1).End(xlUp).Row, 3).Value = countComplete
            wsDashboard.Cells(wsDashboard.Cells(wsDashboard.Rows.Count, 1).End(xlUp).Row, 4).Value = countIncomplete
            wsDashboard.Cells(wsDashboard.Cells(wsDashboard.Rows.Count, 1).End(xlUp).Row, 5).Value = countCorrect
            wsDashboard.Cells(wsDashboard.Cells(wsDashboard.Rows.Count, 1).End(xlUp).Row, 6).Value = countIncorrect
        End If
    Next i
    
    ' Calculate percentages for Incomplete and Incorrect
    Dim lastDashRow As Long
    lastDashRow = wsDashboard.Cells(wsDashboard.Rows.Count, 1).End(xlUp).Row
    For i = 2 To lastDashRow
        If wsDashboard.Cells(i, 2).Value > 0 Then
            wsDashboard.Cells(i, 7).Value = wsDashboard.Cells(i, 4).Value / wsDashboard.Cells(i, 2).Value
            wsDashboard.Cells(i, 8).Value = wsDashboard.Cells(i, 6).Value / wsDashboard.Cells(i, 2).Value
        End If
    Next i
    
    MsgBox "Dashboard updated successfully!"
End Sub
