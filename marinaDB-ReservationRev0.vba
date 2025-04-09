'==========[ Note this is the 1st version, the next version has more protections for doublebookings ]==========
'==========[ For the reservation development process, check box logic portion is not below ]==========
'==========[jas.digital.tools, Jason Santiago 04-09-2025]========== 
'==========[ !!!!! The sub procedure work well as is but will overwrite on a same day rental - I will add the logic after further use of below ]==========
' Form Code
Option Explicit
' Reference to the selected sheet
Public wsRef As Worksheet

Private Sub cmdAddReservation_Click()
    Dim dayRow As Long
    Dim groupName As String
    Dim slipRange As String
    Dim ts As String
    Dim existingSlips As String
    Dim newSlipList As Collection
    Dim slip As Variant
    
    ' ?? Ensure month is selected
    If wsRef Is Nothing Then
        MsgBox "?? Please select a month first using the Set Reference button.", vbExclamation
        Exit Sub
    End If

    ' ??? Get the day of the month
    dayRow = val(InputBox("Enter the day of the month (1–31):", "Which day?"))
    If dayRow < 1 Or dayRow > 31 Then
        MsgBox "? Invalid day selected.", vbCritical
        Exit Sub
    End If

    ' ?? Group name
    groupName = InputBox("Enter group or person name for reservation:", "Group")
    If Trim(groupName) = "" Then Exit Sub

    ' ? Slip Range
    slipRange = InputBox("Enter slip number(s) or range (e.g. 1,6 or 32-34):", "Slips Reserved")
    If Trim(slipRange) = "" Then Exit Sub

    ' ? Parse the slip range into a collection of individual slip numbers
    Set newSlipList = ParseSlipRange(slipRange)
    
    ' ?? Check if any slip already exists in this day’s row
    existingSlips = wsRef.Cells(dayRow, 3).Value ' Column C is where slips are stored

    If HasOverlap(existingSlips, newSlipList) Then
        MsgBox "? One or more slips already reserved on that day!", vbCritical, "Double Booking Detected"
        Exit Sub
    End If

    ' ?? Timestamp
    ts = Format(Now, "mm/dd/yyyy hh:mm AM/PM")

    ' ?? Save reservation
    With wsRef
        .Cells(dayRow, 1).Value = dayRow
        .Cells(dayRow, 2).Value = groupName
        .Cells(dayRow, 3).Value = slipRange
        .Cells(dayRow, 4).Value = ts
    End With

    MsgBox "? Reservation added for day " & dayRow & ".", vbInformation
End Sub

Private Sub UserForm_Initialize()
    ' Hide combo box at first
    cmbMonth.Visible = False
    lblFutureDates.Caption = ""

    ' Load month options
    cmbMonth.Clear
    cmbMonth.AddItem "January"
    cmbMonth.AddItem "February"
    cmbMonth.AddItem "March"
    cmbMonth.AddItem "April"
    cmbMonth.AddItem "May"
    cmbMonth.AddItem "June"
    cmbMonth.AddItem "July"
    cmbMonth.AddItem "August"
    cmbMonth.AddItem "September"
    cmbMonth.AddItem "October"
    cmbMonth.AddItem "November"
    cmbMonth.AddItem "December"
End Sub

Private Sub cmdSetReference_Click()
    cmbMonth.Visible = True
    MsgBox "?? Please choose the reservation month from the dropdown.", vbInformation
End Sub

Private Sub cmbMonth_Change()
    Dim selectedMonth As String
    selectedMonth = cmbMonth.Value

    ' Check if the selected sheet exists
    On Error Resume Next
    Set wsRef = ThisWorkbook.Sheets(selectedMonth)
    On Error GoTo 0

    If wsRef Is Nothing Then
        MsgBox "? Sheet '" & selectedMonth & "' not found.", vbCritical
    Else
        lblFutureDates.Caption = "?? Current Sheet: " & selectedMonth
        MsgBox "? Reference sheet set to " & selectedMonth, vbInformation
    End If
End Sub
'==========[Module Code]==========
'==========[jas.digital.tools, Jason Santiago 04-09-2025]==========                
'Reminder the bounds checking and overwitting probability is not below - saved private
Option Explicit
Public wsRef As Worksheet ' Reference to the selected month sheet
Sub CreateMonthlySheets()
    Dim months As Variant
    Dim i As Integer
    Dim ws As Worksheet
    Dim monthName As String
    Dim sheetExists As Boolean

    months = Array("January", "February", "March", "April", "May", "June", _
                   "July", "August", "September", "October", "November", "December")

    For i = LBound(months) To UBound(months)
        monthName = months(i)
        sheetExists = False

        For Each ws In ThisWorkbook.Sheets
            If ws.Name = monthName Then
                sheetExists = True
                Exit For
            End If
        Next ws

        If Not sheetExists Then
            ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count)).Name = monthName
        End If
    Next i

    MsgBox "Month sheets are ready!", vbInformation
End Sub

' === Reservation Add Button Handler ===
Sub AddReservationWithProtection()
    Dim dayRow As Long
    Dim groupName As String
    Dim slipRange As String
    Dim ts As String
    Dim existingSlips As String
    Dim newSlipList As Collection

    ' ?? Make sure reference sheet is set
    If wsRef Is Nothing Then
        MsgBox "?? Please select a month first using the Set Reference button.", vbExclamation
        Exit Sub
    End If

    ' ?? Prompt for day of the month
    dayRow = val(InputBox("Enter the day of the month (1–31):", "Which day?"))
    If dayRow < 1 Or dayRow > 31 Then
        MsgBox "? Invalid day selected.", vbCritical
        Exit Sub
    End If

    ' ?? Ask for group name
    groupName = InputBox("Enter group or person name for reservation:", "Group")
    If Trim(groupName) = "" Then Exit Sub

    ' ? Ask for slip(s) reserved
    slipRange = InputBox("Enter slip number(s) or range (e.g. 4,5,6 or 10-12):", "Slips Reserved")
    If Trim(slipRange) = "" Then Exit Sub

    ' ?? Parse the input into list of slips
    Set newSlipList = ParseSlipRange(slipRange)

    ' ?? Check for double booking
    existingSlips = wsRef.Cells(dayRow, 3).Value ' Column C is slip log
    If HasOverlap(existingSlips, newSlipList) Then
        MsgBox "? One or more slips already reserved for day " & dayRow & "!", vbCritical
        Exit Sub
    End If

    ' ?? Timestamp
    ts = Format(Now, "mm/dd/yyyy hh:mm AM/PM")

    ' ? Log the reservation
    With wsRef
        .Cells(dayRow, 1).Value = dayRow                 ' Column A: Day
        .Cells(dayRow, 2).Value = groupName             ' Column B: Group
        .Cells(dayRow, 3).Value = slipRange             ' Column C: Slip(s)
        .Cells(dayRow, 4).Value = ts                    ' Column D: Timestamp
    End With

    MsgBox "? Reservation added for day " & dayRow & "!", vbInformation
End Sub

' === Helper to Parse slip range like 1,2,4-6 ===
Function ParseSlipRange(inputStr As String) As Collection
    Dim slips As New Collection
    Dim parts() As String
    Dim p As Variant, startVal As Long, endVal As Long, i As Long

    inputStr = Replace(inputStr, " ", "")
    parts = Split(inputStr, ",")

    For Each p In parts
        If InStr(p, "-") > 0 Then
            startVal = CLng(Split(p, "-")(0))
            endVal = CLng(Split(p, "-")(1))
            For i = startVal To endVal
                slips.Add i
            Next i
        Else
            slips.Add CLng(p)
        End If
    Next

    Set ParseSlipRange = slips
End Function

' === Check if new slips overlap with existing ones ===
Function HasOverlap(existing As String, newSlips As Collection) As Boolean
    Dim existingList As Collection
    Dim s As Variant
    
    If Trim(existing) = "" Then
        Set existingList = New Collection ' empty = no overlap
    Else
        Set existingList = ParseSlipRange(existing)
    End If

    For Each s In newSlips
        If CollectionContains(existingList, s) Then
            HasOverlap = True
            Exit Function
        End If
    Next
    
    HasOverlap = False
End Function

' === Collection lookup ===
Function CollectionContains(coll As Collection, val As Variant) As Boolean
    Dim item As Variant
    For Each item In coll
        If item = val Then
            CollectionContains = True
            Exit Function
        End If
    Next
    CollectionContains = False
End Function





