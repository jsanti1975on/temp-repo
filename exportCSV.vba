Sub ExportStockEntryCSV()

    Dim ws As Worksheet
    Dim exportPath As String
    Dim exportFileName As String
    Dim lastRow As Long
    Dim exportWB As Workbook
    Dim exportWS As Worksheet
    
    ' Reference your output sheet
    Set ws = ThisWorkbook.Sheets("output")
    
    ' Determine the last row of data
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    
    ' Create a new workbook for export
    Set exportWB = Workbooks.Add
    Set exportWS = exportWB.Sheets(1)
    exportWS.Name = "ExportForStockEntry"
    
    ' Copy headers and key columns: SKU (A), Total Units (C), Avg Unit Cost (E)
    exportWS.Cells(1, 1).Value = "Stock Code"
    exportWS.Cells(1, 2).Value = "Qty Recvd"
    exportWS.Cells(1, 3).Value = "Cost Each"

    Dim i As Long, destRow As Long
    destRow = 2

    For i = 2 To lastRow
        exportWS.Cells(destRow, 1).Value = ws.Cells(i, 1).Value ' SKU
        exportWS.Cells(destRow, 2).Value = ws.Cells(i, 3).Value ' Total Units
        exportWS.Cells(destRow, 3).Value = ws.Cells(i, 5).Value ' Avg Unit Cost
        destRow = destRow + 1
    Next i

    ' Save as CSV
    exportPath = ThisWorkbook.Path & "\"
    exportFileName = "StockEntry_" & Format(Now, "yyyymmdd_HHMMSS") & ".csv"

    Application.DisplayAlerts = False
    exportWB.SaveAs fileName:=exportPath & exportFileName, FileFormat:=xlCSV
    exportWB.Close SaveChanges:=False
    Application.DisplayAlerts = True

    MsgBox "Export complete: " & exportFileName, vbInformation

End Sub
