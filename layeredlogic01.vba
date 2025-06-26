Private Sub AddOvernightOrFollowUp()
    Dim i As Integer
    Dim chkBoxName As String
    Dim selectedSlipID As Long
    Dim ws As Worksheet
    Dim newNote As String
    Dim statusChoice As String
    Dim noticeSentDate As Variant
    Dim openDate As Date
    Dim saveTerminationDate As Boolean: saveTerminationDate = False
    
    ' Column assignments
    Dim timestampCol As Long: timestampCol = 9
    Dim userCol As Long: userCol = 10
    Dim noteCol As Long: noteCol = 11
    Dim terminationCol As Long: terminationCol = 12 ' Column L: Date slip will reopen

    Set ws = ThisWorkbook.Sheets("B's-List")
    
    ' === Step 1: Ask about Overnight ===
    Dim response As VbMsgBoxResult
    response = MsgBox("Would you like to mark the slips as Overnight?", vbYesNoCancel + vbQuestion, "Slip Status")

    If response = vbCancel Then Exit Sub

    If response = vbYes Then
        statusChoice = "Overnight"
    Else
        ' === Step 2: Ask about Termination ===
        Dim terminationPrompt As VbMsgBoxResult
        terminationPrompt = MsgBox("Will this be a Notice of Termination?", vbYesNo + vbQuestion, "Notice of Termination")

        If terminationPrompt = vbYes Then
            ' === Step 3: Get the notice sent date ===
            noticeSentDate = InputBox("When was the Notice of Termination sent to Billing?" & vbCrLf & _
                                      "Enter date like MM/DD/YYYY", "Notice Sent Date")
            
            If Not IsDate(noticeSentDate) Then
                MsgBox "Invalid date format. Operation cancelled.", vbExclamation
                Exit Sub
            End If
            
            If CDate(noticeSentDate) > Date Then
                MsgBox "The date entered is in the future. Please confirm.", vbExclamation
                Exit Sub
            End If

            ' === Step 4: Calculate 14 days ahead ===
            openDate = CDate(noticeSentDate) + 14
            saveTerminationDate = True
            statusChoice = "Follow-Up"
        Else
            statusChoice = "Follow-Up"
        End If
    End If

    ' === Step 5: Apply to checked slips ===
    For i = 1 To 80
        chkBoxName = "CheckBox" & i
        If Me.Controls(chkBoxName).Value = True Then
            selectedSlipID = i
            
            ' Set status (Column A)
            ws.Cells(selectedSlipID, 1).Value = statusChoice

            ' If termination date applies, save it
            If saveTerminationDate Then
                ws.Cells(selectedSlipID, terminationCol).Value = openDate
            End If

            ' === Step 6: Ask for a note ===
            newNote = InputBox("Enter a note for Slip " & selectedSlipID & " (" & statusChoice & "):", "Add Note")
            
            If Trim(newNote) <> "" Then
                ws.Cells(selectedSlipID, noteCol).Value = newNote
                ws.Cells(selectedSlipID, userCol).Value = Application.UserName
                ws.Cells(selectedSlipID, timestampCol).Value = Now
            End If
        End If
    Next i

    MsgBox "Selected slip(s) have been updated.", vbInformation

    ' Refresh checkbox colors
    ApplyCheckboxColors
End Sub
