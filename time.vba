Sub SetupTimesheetLayout()

    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets("Timesheet")
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Sheets.Add
        ws.Name = "Timesheet"
    Else
        ws.Cells.Clear
    End If
    On Error GoTo 0
    
    Dim headers As Variant
    headers = Array("Date", "Day", "Time In", "Time Out", "Break (hrs)", "Total Hours", "Job Code", "Description")
    
    Dim i As Integer
    For i = 0 To UBound(headers)
        ws.Cells(1, i + 1).Value = headers(i)
        ws.Cells(1, i + 1).Font.Bold = True
        ws.Columns(i + 1).ColumnWidth = 15
        ws.Cells(1, i + 1).Interior.Color = RGB(200, 200, 200)
        ws.Cells(1, i + 1).Borders.LineStyle = xlContinuous
    Next i
    
    ' Add light row borders for a clean grid
    For i = 2 To 32 ' 31 rows for one month
        ws.Range("A" & i & ":H" & i).Borders.LineStyle = xlHairline
    Next i

    ws.Range("A1:H1").HorizontalAlignment = xlCenter
    MsgBox "Timesheet layout ready", vbInformation

End Sub
