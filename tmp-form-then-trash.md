# Objects on the Form (frmTenantCallLog)
```bash
Control Name	Type	Purpose
lblTenantName	Label	Displays the tenant’s name.
cmbTenantList	ComboBox	Dropdown to select a tenant (auto-filled from tenant database).
lblPhoneNumber	Label	Displays the phone number of the selected tenant.
txtPhoneNumber	TextBox	Displays the phone number (non-editable).
lblStatus	Label	Label for call status options.
optSpokeToTenant	OptionButton	Select if the specialist spoke to the tenant.
optLeftMessage	OptionButton	Select if the specialist left a message.
lblNotes	Label	Label for txtNotes.
txtNotes	TextBox	Small note about the call (only if tenant was reached).
cmdLogCall	CommandButton	Button to log the call entry into CallLog.
cmdCancel	CommandButton	Button to close the form without saving.
  ```


# 🔹 Non-Repudiation Strategy
```Bash
To ensure non-repudiation, we log the following details:

Timestamp (callTime): Logs the date and time of the call.
Caller (Application.UserName) (Optional Enhancement): Captures the Excel username to track who logged the call.
Preventing Unauthorized Changes:
The CallLog sheet can be password-protected to prevent users from modifying call records after entry.
Digital Signatures can be used if needed.
To add a username column:
```


```vba
.Cells(iRow, 6).Value = Application.UserName ' Logs the specialist's username
```
