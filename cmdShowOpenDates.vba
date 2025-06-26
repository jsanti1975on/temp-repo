' 06252025 working sub to write
' Notice of Termination sub process was written at home and needs to be inserted before write of below.
'----
'==== Objects ====
' lstTerminations
' cmdShowOpenDates
'==== Objects ====

Private Sub cmdShowOpenDates_Click()
    Dim ws As Worksheet
    Dim i As Long
    Dim slipID As Long
    Dim openDate As Variant
    Dim resultList As Collection
    Dim item As Variant

    Set ws = ThisWorkbook.Sheets("B's-List")
    Set resultList = New Collection

    ' Clear previous listbox entries
    lstTerminations.Clear

    ' Loop through slips 1–80
    For i = 1 To 80
        openDate = ws.Cells(i, 12).Value ' Column L = termination/open date

        ' Only include rows where a valid date is set
        If IsDate(openDate) Then
            slipID = ws.Cells(i, 2).Value ' Column B = Slip Number
            resultList.Add Array(slipID, Format(openDate, "mm/dd/yyyy"))
        End If
    Next i

    ' Populate the ListBox
    If resultList.Count > 0 Then
        lstTerminations.ColumnCount = 2
        lstTerminations.Clear

        For Each item In resultList
            lstTerminations.AddItem item(0)
            lstTerminations.List(lstTerminations.ListCount - 1, 1) = item(1)
        Next item
    Else
        MsgBox "No termination dates found.", vbInformation
    End If
End Sub
