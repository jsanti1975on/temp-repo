Option Explicit
' 04-02-2025 Update by Jason Santiago to create a note taking method to the visual marina map 
Private Sub cmdClear_Click()
    ClearAllCheckboxes
End Sub

Private Sub cmdUpdateTestSheet_Click()
    AddOvernightOrFollowUp
End Sub

Private Sub ClearAllCheckboxes()
    Dim i As Integer
    Dim chkBoxName As String

    For i = 1 To 80
        chkBoxName = "CheckBox" & i
        If Not Me.Controls(chkBoxName) Is Nothing Then
            Me.Controls(chkBoxName).Value = False
            Me.Controls(chkBoxName).BackColor = &H8000000F ' Default
        End If
    Next i
End Sub

Private Sub UserForm_Initialize()
    Dim ws As Worksheet
    Dim i As Integer
    Dim chkBoxName As String
    Dim cellValue As String

    Set ws = ThisWorkbook.Sheets("ParsedData")
    lblDetails.Caption = "Click a yellow or blue checkbox to view details."

    For i = 1 To 80
        chkBoxName = "CheckBox" & i
        If Not Me.Controls(chkBoxName) Is Nothing Then
            cellValue = Trim(ws.Cells(i, 1).Value)

            With Me.Controls(chkBoxName)
                Select Case True
                    Case cellValue = "Open_Slip"
                        .Value = False
                        .BackColor = RGB(0, 255, 0) ' Green
                    Case cellValue = "COMMERCIAL"
                        .Value = True
                        .BackColor = RGB(192, 192, 192) ' Grey
                    Case cellValue = "Overnight"
                        .Value = True
                        .BackColor = RGB(255, 255, 0) ' Yellow
                    Case cellValue = "Follow-Up"
                        .Value = True
                        .BackColor = RGB(0, 0, 255) ' Blue
                    Case InStr(cellValue, ",") > 0
                        .Value = True
                        .BackColor = RGB(255, 0, 0) ' Red
                    Case Else
                        .Value = False
                        .BackColor = &H8000000F ' Default
                End Select
            End With
        End If
    Next i
End Sub

Private Sub AddOvernightOrFollowUp()
    Dim i As Integer
    Dim chkBoxName As String
    Dim selectedSlipID As Integer
    Dim userChoice As VbMsgBoxResult
    Dim ws As Worksheet
    Dim newNote As String
    Dim timestampCol As Long: timestampCol = 9
    Dim userCol As Long: userCol = 10
    Dim noteCol As Long: noteCol = 11

    Set ws = ThisWorkbook.Sheets("ParsedData")

    userChoice = MsgBox("Would you like to mark selected slips as OVERNIGHT (Yes) or FOLLOW-UP (No)?", _
                        vbYesNoCancel + vbQuestion, "Select Status")
    If userChoice = vbCancel Then Exit Sub

    For i = 1 To 80
        chkBoxName = "CheckBox" & i
        If Me.Controls(chkBoxName).Value = True Then
            selectedSlipID = i

            ' Update status
            Select Case userChoice
                Case vbYes: ws.Cells(selectedSlipID, 1).Value = "Overnight"
                Case vbNo:  ws.Cells(selectedSlipID, 1).Value = "Follow-Up"
            End Select

            ' Prompt for note
            newNote = InputBox("Enter a note for Slip " & selectedSlipID & " (" & ws.Cells(selectedSlipID, 1).Value & "):", "Add Note")
            If Trim(newNote) <> "" Then
                ws.Cells(selectedSlipID, noteCol).Value = newNote
                ws.Cells(selectedSlipID, userCol).Value = Application.UserName
                ws.Cells(selectedSlipID, timestampCol).Value = Now
            End If
        End If
    Next i

    MsgBox "Selected slips have been updated.", vbInformation
    ApplyCheckboxColors
End Sub

Private Sub ApplyCheckboxColors()
    Dim ws As Worksheet
    Dim i As Integer
    Dim chkBoxName As String
    Dim cellValue As String

    Set ws = ThisWorkbook.Sheets("ParsedData")

    For i = 1 To 80
        chkBoxName = "CheckBox" & i
        If Not Me.Controls(chkBoxName) Is Nothing Then
            cellValue = Trim(ws.Cells(i, 1).Value)

            With Me.Controls(chkBoxName)
                Select Case True
                    Case cellValue = "Open_Slip"
                        .Value = False
                        .BackColor = RGB(0, 255, 0)
                    Case cellValue = "COMMERCIAL"
                        .Value = True
                        .BackColor = RGB(192, 192, 192)
                    Case cellValue = "Overnight"
                        .Value = True
                        .BackColor = RGB(255, 255, 0)
                    Case cellValue = "Follow-Up"
                        .Value = True
                        .BackColor = RGB(0, 0, 255)
                    Case InStr(cellValue, ",") > 0
                        .Value = True
                        .BackColor = RGB(255, 0, 0)
                    Case Else
                        .Value = False
                        .BackColor = &H8000000F
                End Select
            End With
        End If
    Next i
End Sub

Private Sub HandleCheckboxClick(index As Long)
    Dim ws As Worksheet
    Dim rowNum As Long
    Dim note As String, user As String
    Dim timestamp As Variant
    Dim cellValue As String

    Set ws = ThisWorkbook.Sheets("ParsedData")
    rowNum = index
    cellValue = Trim(ws.Cells(rowNum, 1).Value)

    If cellValue = "Overnight" Or cellValue = "Follow-Up" Then
        note = ws.Cells(rowNum, 11).Value
        user = ws.Cells(rowNum, 10).Value
        timestamp = ws.Cells(rowNum, 9).Value

        If IsDate(timestamp) Then timestamp = Format(timestamp, "mm/dd/yyyy hh:mm AM/PM")

        lblDetails.Caption = "Slip " & index & " (" & cellValue & ")" & vbCrLf & _
                             "Note: " & note & vbCrLf & _
                             "By: " & user & vbCrLf & _
                             "At: " & timestamp
    Else
        lblDetails.Caption = "Slip " & index & " is not marked as Overnight or Follow-Up."
    End If
End Sub

'====[One-Liner to create the 80 handlers ]====
For i = 1 To 80: Debug.Print "Private Sub CheckBox" & i & "_Click(): HandleCheckboxClick " & i & ": End Sub": Next i
    







      
