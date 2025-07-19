Sub HighlightFieldsForStockEntry()

    Dim ws As Worksheet
    Dim lastRow As Long
    
    ' Set your output sheet
    Set ws = ThisWorkbook.Sheets("output")
    
    ' Find last row with data in Column A
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row

    ' Clear previous formatting
    ws.Range("A1:E" & lastRow).Interior.ColorIndex = xlNone
    ws.Range("A1:E1").Font.Bold = False

    ' Highlight necessary fields
    With ws
        ' SKU (Stock Code) - Column A
        .Range("A1:A" & lastRow).Interior.Color = RGB(0, 176, 80) ' Green
        .Range("A1").Font.Bold = True
        
        ' Total Units - Column C
        .Range("C1:C" & lastRow).Interior.Color = RGB(0, 176, 80)
        .Range("C1").Font.Bold = True

        ' Avg Unit Cost - Column E
        .Range("E1:E" & lastRow).Interior.Color = RGB(0, 176, 80)
        .Range("E1").Font.Bold = True
    End With

    MsgBox "Highlighted fields for Retail System Input!", vbInformation

End Sub
