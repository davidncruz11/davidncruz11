Sub UpdateSummary()
    Dim wsCharts As Worksheet, wsSummary As Worksheet
    Dim lastRowCharts As Long, lastRowSummary As Long
    Dim i As Long, doc As String
    Dim countRange As Range, existingDoc As Range
    
    Set wsCharts = ThisWorkbook.Sheets("Audited Charts")
    Set wsSummary = ThisWorkbook.Sheets("Summary")
    
    lastRowCharts = wsCharts.Cells(wsCharts.Rows.Count, 1).End(xlUp).Row
    lastRowSummary = wsSummary.Cells(wsSummary.Rows.Count, 1).End(xlUp).Row + 1
    
    ' Loop through the rows in the Audited Charts sheet
    For i = 2 To lastRowCharts
        If wsCharts.Cells(i, 1).Value <> "" Then
            doc = wsCharts.Cells(i, 2).Value ' Document name
            Set countRange = wsSummary.Range("A2:A" & lastRowSummary - 1) ' Range to search for the document in the summary
            
            ' Check if the document already exists in the summary
            Set existingDoc = countRange.Find(doc, LookIn:=xlValues, LookAt:=xlWhole)
            
            If Not existingDoc Is Nothing Then
                ' Document exists, update the count
                existingDoc.Offset(0, 1).Value = existingDoc.Offset(0, 1).Value + 1 ' Increment count
                existingDoc.Offset(0, 3).Value = Application.CountIf(wsCharts.Range("D2:D" & lastRowCharts), "Yes") ' Correct
                existingDoc.Offset(0, 4).Value = Application.CountIf(wsCharts.Range("D2:D" & lastRowCharts), "No") ' Incorrect
                existingDoc.Offset(0, 6).Formula = "=IF(B" & existingDoc.Row & ">0, D" & existingDoc.Row & "/B" & existingDoc.Row & ", 0)" ' % Incomplete
                existingDoc.Offset(0, 7).Formula = "=IF(B" & existingDoc.Row & ">0, F" & existingDoc.Row & "/B" & existingDoc.Row & ", 0)" ' % Incorrect
            Else
                ' Document does not exist, add a new entry
                wsSummary.Cells(lastRowSummary, 1).Value = doc ' Document name
                wsSummary.Cells(lastRowSummary, 2).Value = 1 ' Initial count
                wsSummary.Cells(lastRowSummary, 3).Value = Application.CountIf(wsCharts.Range("C2:C" & lastRowCharts), "Yes") ' Complete
                wsSummary.Cells(lastRowSummary, 4).Value = Application.CountIf(wsCharts.Range("C2:C" & lastRowCharts), "No") ' Incomplete
                wsSummary.Cells(lastRowSummary, 5).Value = Application.CountIf(wsCharts.Range("D2:D" & lastRowCharts), "Yes") ' Correct
                wsSummary.Cells(lastRowSummary, 6).Value = Application.CountIf(wsCharts.Range("D2:D" & lastRowCharts), "No") ' Incorrect
                wsSummary.Cells(lastRowSummary, 7).Formula = "=IF(B" & lastRowSummary & ">0, D" & lastRowSummary & "/B" & lastRowSummary & ", 0)" ' % Incomplete
                wsSummary.Cells(lastRowSummary, 8).Formula = "=IF(B" & lastRowSummary & ">0, F" & lastRowSummary & "/B" & lastRowSummary & ", 0)" ' % Incorrect
                
                lastRowSummary = lastRowSummary + 1 ' Move to the next row in Summary
            End If
        End If
    Next i

    MsgBox "Summary updated successfully!"
End Sub
