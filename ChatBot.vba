Private Sub cmdLog_Click()
    Dim wsLog As Worksheet
    Dim lastRow As Long
    Dim matchRow As Long
    Dim question As String
    Dim i As Long

    question = Trim(txtQuestion.Value)
    If question = "" Then
        MsgBox "You forgot your question, sailor!", vbExclamation, "Captain Clip"
        Exit Sub
    End If

    On Error Resume Next
    Set wsLog = Sheets("CaptainLog")
    If wsLog Is Nothing Then
        Set wsLog = Sheets.Add(After:=Sheets(Sheets.Count))
        wsLog.Name = "CaptainLog"
        wsLog.Range("A1:D1").Value = Array("Timestamp", "Question", "Count", "Last Updated")
    End If
    On Error GoTo 0

    lastRow = wsLog.Cells(wsLog.Rows.Count, "B").End(xlUp).Row
    matchRow = 0

    For i = 2 To lastRow
        If LCase(wsLog.Cells(i, "B").Value) = LCase(question) Then
            matchRow = i
            Exit For
        End If
    Next i

    If matchRow > 0 Then
        wsLog.Cells(matchRow, "C").Value = wsLog.Cells(matchRow, "C").Value + 1
        wsLog.Cells(matchRow, "D").Value = Now
    Else
        lastRow = lastRow + 1
        wsLog.Cells(lastRow, "A").Value = Now
        wsLog.Cells(lastRow, "B").Value = question
        wsLog.Cells(lastRow, "C").Value = 1
        wsLog.Cells(lastRow, "D").Value = Now
    End If

    MsgBox "Logged, Captain! ☑️", vbInformation, "Captain Clip"
    txtQuestion.Value = ""
    txtQuestion.SetFocus
End Sub