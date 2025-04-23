Private Sub btnStartTask_Click()
    If lblRowIndex.Caption <> "" Then
        MsgBox "You already have an open task. End it first.", vbExclamation
        Exit Sub
    End If

    Dim ws As Worksheet
    Dim taskName As String
    Dim newRow As Long

    Set ws = ThisWorkbook.Sheets("TaskTracker")
    taskName = Trim(txtTaskName.Text)
    
    If taskName = "" Then
        MsgBox "Enter a task name first.", vbExclamation
        Exit Sub
    End If
    
    newRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row + 1
    ws.Cells(newRow, 1).Value = taskName
    ws.Cells(newRow, 2).Value = Now
    ws.Cells(newRow, 2).NumberFormat = "hh:mm AM/PM"
    
    lblRowIndex.Caption = CStr(newRow)
    
    MsgBox "Task '" & taskName & "' started at " & Format(Now, "hh:mm AM/PM")
End Sub
