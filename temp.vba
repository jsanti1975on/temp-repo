Option Explicit

' 🚀 Load Tenant Names & Call Log Data on Form Initialization
Private Sub UserForm_Initialize()
    Dim wsCallLog As Worksheet, wsTenantList As Worksheet
    Dim lastRowCallLog As Long, lastRowTenant As Long
    Dim i As Long
    
    Set wsCallLog = ThisWorkbook.Sheets("CallLog")
    Set wsTenantList = ThisWorkbook.Sheets("TenantList")

    ' Get last row in CallLog and TenantList
    lastRowCallLog = wsCallLog.Cells(wsCallLog.Rows.Count, "A").End(xlUp).Row
    lastRowTenant = wsTenantList.Cells(wsTenantList.Rows.Count, "A").End(xlUp).Row

    ' 🏠 Populate cmbTenantList from TenantList
    cmbTenantList.Clear
    For i = 2 To lastRowTenant
        cmbTenantList.AddItem wsTenantList.Cells(i, 1).Value ' Name
    Next i

    ' 🎛 Configure ListBox: lstTenantNotes
    With lstTenantNotes
        .Clear
        .ColumnCount = 6
        .ColumnWidths = "100,100,100,120,150,100"
    End With

    ' ☎ Configure ListBox: lstCallBack
    With lstCallBack
        .Clear
        .ColumnCount = 6
        .ColumnWidths = "100,100,100,120,150,100"
    End With

    ' 🔍 Load Call Log data into lstTenantNotes
    For i = 2 To lastRowCallLog
        lstTenantNotes.AddItem wsCallLog.Cells(i, 1).Value ' Name
        lstTenantNotes.List(lstTenantNotes.ListCount - 1, 1) = wsCallLog.Cells(i, 2).Value ' Phone
        lstTenantNotes.List(lstTenantNotes.ListCount - 1, 2) = wsCallLog.Cells(i, 3).Value ' Contacted
        lstTenantNotes.List(lstTenantNotes.ListCount - 1, 3) = wsCallLog.Cells(i, 4).Value ' Timestamp
        lstTenantNotes.List(lstTenantNotes.ListCount - 1, 4) = wsCallLog.Cells(i, 5).Value ' Notes
        lstTenantNotes.List(lstTenantNotes.ListCount - 1, 5) = wsCallLog.Cells(i, 6).Value ' User
        
        ' 📌 Load into lstCallback if follow-up is needed
        If wsCallLog.Cells(i, 3).Value = "Left Message" Or _
           wsCallLog.Cells(i, 3).Value = "No Answer" Or _
           wsCallLog.Cells(i, 3).Value = "Bad Phone Number" Or _
           wsCallLog.Cells(i, 3).Value = "Tenant Requested Callback" Then

            lstCallBack.AddItem wsCallLog.Cells(i, 1).Value
            lstCallBack.List(lstCallBack.ListCount - 1, 1) = wsCallLog.Cells(i, 2).Value
            lstCallBack.List(lstCallBack.ListCount - 1, 2) = wsCallLog.Cells(i, 3).Value
            lstCallBack.List(lstCallBack.ListCount - 1, 3) = wsCallLog.Cells(i, 4).Value
            lstCallBack.List(lstCallBack.ListCount - 1, 4) = wsCallLog.Cells(i, 5).Value
            lstCallBack.List(lstCallBack.ListCount - 1, 5) = wsCallLog.Cells(i, 6).Value
        End If
    Next i

    ' 🎯 Populate Call Status dropdown (Restricted to List)
    With cmbReason
        .Clear
        .Style = fmStyleDropDownList ' Prevents typing custom text
        .AddItem "Spoke to Tenant"
        .AddItem "Left Message"
        .AddItem "No Answer"
        .AddItem "Bad Phone Number"
        .AddItem "Tenant Requested Callback"
        .AddItem "Confirmed Compliance"
        .AddItem "Refused to Move Boat"
    End With
End Sub

' 📞 Auto-Fill Phone Number When Tenant is Selected
Private Sub cmbTenantList_Change()
    Dim ws As Worksheet
    Dim rng As Range
    Dim cell As Range

    Set ws = ThisWorkbook.Sheets("TenantList")
    Set rng = ws.Range("A2:B" & ws.Cells(ws.Rows.Count, "A").End(xlUp).Row) ' Column B = Phone Numbers

    ' Search for selected tenant's phone number
    For Each cell In rng
        If cell.Value = cmbTenantList.Value Then
            txtPhoneNumber.Value = cell.Offset(0, 1).Value ' Get phone number from column B
            Exit For
        End If
    Next cell
End Sub

' 📝 Handle Click Event in lstTenantNotes (Display Notes)
Private Sub lstTenantNotes_Click()
    Dim selectedRow As Integer
    Dim noteText As String
    
    ' Ensure a selection is made
    If lstTenantNotes.ListIndex = -1 Then Exit Sub
    
    ' Retrieve the note from Column 5 (E)
    selectedRow = lstTenantNotes.ListIndex
    noteText = lstTenantNotes.List(selectedRow, 4) ' Column 4 (0-based) contains Notes
    
    ' Display the note
    If Trim(noteText) <> "" Then
        MsgBox "Note: " & vbCrLf & noteText, vbInformation, "Tenant Note Details"
    Else
        MsgBox "No note available for this entry.", vbExclamation, "No Note"
    End If
End Sub

' 🚀 Log Call to `CallLog` Sheet
Private Sub cmdLogCall_Click()
    Dim ws As Worksheet
    Dim iRow As Long
    Dim callStatus As String
    Dim callTime As String
    
    ' Ensure a tenant is selected
    If cmbTenantList.Value = "" Then
        MsgBox "Please select a tenant.", vbExclamation, "Missing Data"
        Exit Sub
    End If
    
    ' Ensure a reason is selected
    If cmbReason.Value = "" Then
        MsgBox "Please select a reason for the call.", vbExclamation, "Missing Data"
        Exit Sub
    End If

    ' Capture call status
    callStatus = cmbReason.Value
    
    ' Disable Notes if "Left Message" is selected
    If callStatus = "Left Message" Then
        txtNotes.Value = ""
        txtNotes.Enabled = False
    Else
        txtNotes.Enabled = True
    End If

    ' Capture current time
    callTime = Format(Now(), "MM/DD/YYYY HH:MM AM/PM")

    ' Set reference to Call Log worksheet
    Set ws = ThisWorkbook.Sheets("CallLog")
    
    ' Find the next available row
    iRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row + 1
    
    ' Log the call
    With ws
        .Cells(iRow, 1).Value = cmbTenantList.Value ' Tenant Name
        .Cells(iRow, 2).Value = txtPhoneNumber.Value ' Phone Number
        .Cells(iRow, 3).Value = callStatus ' Call Status
        .Cells(iRow, 4).Value = callTime ' Timestamp
        .Cells(iRow, 5).Value = txtNotes.Value ' Notes
        .Cells(iRow, 6).Value = Application.UserName ' Logs the specialist's username
    End With

    ' Confirmation message
    MsgBox "Call logged successfully!", vbInformation, "Success"

    ' Reset the form
    cmbTenantList.Value = ""
    txtPhoneNumber.Value = ""
    txtNotes.Value = ""
    cmbReason.Value = ""
    
    ' Reload form to update Call Lists
    UserForm_Initialize
End Sub

' ❌ Close Form
Private Sub cmdCancel_Click()
    Unload Me
End Sub
