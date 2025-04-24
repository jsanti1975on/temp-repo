# TaskTrackerForm 

```bash
txtTaskName
txtNotes | Notes for task (multiline = True, scrollbars = 2 (vertical)).
lblTotalTime | Display total time spent. Set Caption = "Total Time Spent: 0 hours".
lblRowIndex | Hidden. Stores row index. Set Visible = False.
txtActiveTask | Read-only or hidden. Shows resumed task name. Set Visible = False.
btnStartTask | 
btnEndTask
btnReset
```

## Adding Notes

```plaintext
Adding notes
txtNotes:

MultiLine = True

EnterKeyBehaviour = True

ScrollBars = fmScrollBarsVertical

txtActiveTask:

Locked = True if visible

BackColor light gray (to show it's non-editable)
```

```bash
+--------------------------+
| [ txtTaskName        ]   |
| [ txtNotes           ]   |
| [ lblTotalTime       ]   |
| [ btnStartTask ]          [ btnEndTask ]    [ btnReset ] |
| [ txtActiveTask ]         (hidden or read-only)           |
| [ lblRowIndex   ]         (hidden)                        |
+--------------------------+
```

```vba
Sub GenerateDailySummary()
    Dim srcWS As Worksheet, destWS As Worksheet
    Dim lastRow As Long, summaryRow As Long
    Dim taskDict As Object, r As Long
    Dim taskDate As String, totalTime As Double

    Set srcWS = ThisWorkbook.Sheets("TaskTracker")
    Set taskDict = CreateObject("Scripting.Dictionary")
    
    lastRow = srcWS.Cells(srcWS.Rows.Count, 2).End(xlUp).Row

    ' Aggregate total time by start date (rounded to day)
    For r = 2 To lastRow
        If IsDate(srcWS.Cells(r, 2).Value) Then
            taskDate = Format(srcWS.Cells(r, 2).Value, "mm/dd/yyyy")
            totalTime = srcWS.Cells(r, 5).Value
            If taskDict.exists(taskDate) Then
                taskDict(taskDate) = taskDict(taskDate) + totalTime
            Else
                taskDict.Add taskDate, totalTime
            End If
        End If
    Next r

    ' Output to new sheet
    On Error Resume Next
    Application.DisplayAlerts = False
    Worksheets("DailySummary").Delete
    Application.DisplayAlerts = True
    On Error GoTo 0

    Set destWS = ThisWorkbook.Sheets.Add
    destWS.Name = "DailySummary"
    destWS.Range("A1:B1").Value = Array("Date", "Total Time (hrs)")
    
    summaryRow = 2
    Dim key
    For Each key In taskDict.Keys
        destWS.Cells(summaryRow, 1).Value = key
        destWS.Cells(summaryRow, 2).Value = Round(taskDict(key), 2)
        summaryRow = summaryRow + 1
    Next key
    
    destWS.Columns("A:B").AutoFit
    MsgBox "Daily summary generated."
End Sub
```

```vba
' Call this weekly, daily, or with a button for manual backups.
Sub BackupTaskLog()
    Dim srcWS As Worksheet, backupWS As Worksheet
    Dim backupName As String
    Dim timestamp As String

    Set srcWS = ThisWorkbook.Sheets("TaskTracker")
    timestamp = Format(Now, "yyyymmdd_hhmmss")
    backupName = "Backup_" & timestamp

    srcWS.Copy After:=Sheets(Sheets.Count)
    ActiveSheet.Name = backupName

    MsgBox "Backup created: " & backupName
End Sub
```


