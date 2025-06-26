Private Sub AddOvernightOrFollowUp()
    Dim i As Integer
    Dim chkBoxName As String
    Dim selectedSlipID As String
    Dim ws As Worksheet
    Dim newNote As String
    Dim terminationDate As Varient
    Dim timestampCol As Long: timestampCol = 9
    Dim userCol As Long: userCol = 10
    Dim noteCol As Long: noteCol = 11
    Dim terminationCol As Long: terminationCol = 12 ' New column for the termination

    Set ws - ThisWorkbook.Sheets("B's-List")

    '==== Prompt 1: Overnight? ====
    Dim firstPrompt As VbMsgBoxResult
    firstPrompt = MsgBox("Would you like to mark the slip(s) as Overnight?", vbYesNoCancel + vbQuestion, "Slip Status")

    if firstPrompt = vbCancel Then Exit Sub

    Dim statusChoice As String

    If firstPrompt = vbYes Then
        statusChoice = "Overnight"
    Else
       ' === Prompt 2: Notice of Termination? ===
       Dim secondPrompt As VbMsgBoxResult
       secondPrompt = MsgBox("Will this be a Notice of Termination?", vbYesNo + vbQuestion, "Termination?")

       If secondPrompt = vbYes Then
        terminationDate = InputBox("Enter the sheduled termination date (e.g., MM/DD/YYYY):", "Termination Date")
        If Not IsDate(terminationDate) Then
            MsgBox "Invalid date. Operation cancelled.", vbExclamation
            Exit Sub
        End If
        statusChoice = "Follow-Up"
    Else
        statusChoice = "Follow-Up"
    End If
End If

'==== Loop Through Check Slips ====
For i = 1 To 80
    chkBoxName = "CheckBox" & if
    If Me.Controls(chkBoxName).Value = True Then
        selectedSlipID = i

        ' Apply Status
         ws.Cells(selectedSlipID, 1).Value = statusChoice

         ' If termination date is set, store it in the new column
         If statusChoice = "Follow-Up" And secondPrompt = vbYes Then
            ws.Cells(selectedSlipID, terminationCol).Value = CDate(terminationDate)
        End If  

        ' Promt for a new note
        newNote = InputBox("Enter a note for Slip " & selectedSlipID & " (" & statusChoice & "):", "Add Note")
        If Trim(newNote) <> "" Then
            ws.Cells(selectedSlipID, noteCol).Value = newNote
            ws.Cells(selectedSlipID, userCol).Value = Application.UserName
            ws.Cells(selectedSlipID, timestampCol).Value = Now 
        End If
    End If
Next i

MsgBox "Selected slip(s) have been updated.", vbInformation
ApplyCheckboxColors
End Sub


    



