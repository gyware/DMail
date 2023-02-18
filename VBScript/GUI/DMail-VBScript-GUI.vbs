recipient = InputBox("To", "DMail (VBScript)") 'To
Set re = New RegExp
With re
	.Pattern = "([-!#-'*+/-9=?A-Z^-~]+(\.[-!#-'*+/-9=?A-Z^-~]+)*|""([]!#-[^-~ \t]|(\\[\t -~]))+"")@([-!#-'*+/-9=?A-Z^-~]+(\.[-!#-'*+/-9=?A-Z^-~]+)*|\[[\t -Z^-~]*])"
End With
Do Until success = True
	If re.Test(recipient) = True Then
		success = True
	Else
		If IsEmpty(recipient) Then
			button = MsgBox("""To"" is required.", vbCritical + vbAbortRetryIgnore + vbDefaultButton2 + vbSystemModal, "DMail (VBScript)")
			Select Case button
			Case vbAbort
				WScript.Quit
			Case vbRetry
				success = False
				recipient = InputBox("To", "DMail (VBScript)")
			Case vbIgnore
				button = MsgBox("Are you sure you want to ignore this error? Ignoring may cause unexpected problems.", vbExclamation + vbYesNo + vbDefaultButton2 + vbSystemModal, "DMail (VBScript)")
				Select Case button
				Case vbYes
					success = True
				Case vbNo
					success = False
				End Select
			End Select
		Else
			button = MsgBox("""" & recipient & """ is not a valid email address.", vbCritical + vbAbortRetryIgnore + vbDefaultButton2 + vbSystemModal, "DMail (VBScript)")
			Select Case button
			Case vbAbort
				WScript.Quit
			Case vbRetry
				success = False
				recipient = InputBox("To", "DMail (VBScript)", recipient)
			Case vbIgnore
				button = MsgBox("Are you sure you want to ignore this error? Ignoring may cause unexpected problems.", vbExclamation + vbYesNo + vbDefaultButton2 + vbSystemModal, "DMail (VBScript)")
				Select Case button
				Case vbYes
					success = True
				Case vbNo
					success = False
				End Select
			End Select
		End If
	End If
Loop
success = False

From = InputBox("From", "DMail (VBScript)") 'From
Do Until success = True
	If re.Test(from) = True Then
		success = True
	Else
		If IsEmpty(from) Then
			button = MsgBox("""From"" is required.", vbCritical + vbAbortRetryIgnore + vbDefaultButton2 + vbSystemModal, "DMail (VBScript)")
			Select Case button
			Case vbAbort
				WScript.Quit
			Case vbRetry
				success = False
				from = InputBox("From", "DMail (VBScript)")
			Case vbIgnore
				button = MsgBox("Are you sure you want to ignore this error? Ignoring may cause unexpected problems.", vbExclamation + vbYesNo + vbDefaultButton2 + vbSystemModal, "DMail (VBScript)")
				Select Case button
				Case vbYes
					success = True
				Case vbNo
					success = False
				End Select
			End Select
		Else
			button = MsgBox("""" & from & """ is not a valid email address.", vbCritical + vbAbortRetryIgnore + vbDefaultButton2 + vbSystemModal, "DMail (VBScript)")
			Select Case button
			Case vbAbort
				WScript.Quit
			Case vbRetry
				success = False
				From = InputBox("From", "DMail (VBScript)", from)
			Case vbIgnore
				button = MsgBox("Are you sure you want to ignore this error? Ignoring may cause unexpected problems.", vbExclamation + vbYesNo + vbDefaultButton2 + vbSystemModal, "DMail (VBScript)")
				Select Case button
				Case vbYes
					success = True
				Case vbNo
					success = False
				End Select
			End Select
		End If
	End If
Loop
success = False		
Set re = Nothing

Set mail = CreateObject("CDO.Message")

mail.To = recipient

mail.From = From

mail.Subject = InputBox("Subject (optional)", "DMail (VBScript)") 'Subject

mail.TextBody = InputBox("Body (optional)", "DMail (VBScript)") 'Body

customSMTP = InputBox("SMTP Server (optional)", "DMail (VBScript)") 'Custom SMTP Server

If customSMTP = "" Then
	mail.Configuration.Fields.Item ("http://schemas.microsoft.com/cdo/configuration/sendusing") = 2
	
	On Error Resume Next
	
	Do Until success = True
		mailDomain = Split(recipient, "@")(1)
		If Err.Number <> 0 Then
			button = MsgBox("Error: Domain extraction failed." & vbCrLf & "To: " & recipient & vbCrLf & "From: " & from & vbCrLf & "Description: " & Err.Description & vbCrLf & "Note: If you have pressed ""Ignore"" when entering the ""To"" address, you are ignoring that the email address is not valid and is hence causing this error.", vbCritical + vbAbortRetryIgnore + vbDefaultButton2 + vbSystemModal, "DMail (VBScript)")
			Select Case button
			Case vbAbort
				WScript.Quit
			Case vbRetry
				success = False
				recipient = InputBox("To", "DMail (VBScript)", recipient)
				mailDomain = Split(recipient, "@")(1)
			Case vbIgnore
				button = MsgBox("Are you sure you want to ignore this error? Ignoring may cause unexpected problems.", vbExclamation + vbYesNo + vbDefaultButton2 + vbSystemModal, "DMail (VBScript)")
				Select Case button
				Case vbYes
					success = True
				Case vbNo
					success = False
				End Select
			End Select
		Else
			success = True
		End If
	Loop
	success = False
	
	Set oShell = CreateObject("WScript.Shell")
	Do Until success = True
		If mailDomain <> "" Then
			Set sOutput = oShell.Exec("nslookup -q=mx " & mailDomain)
			sOutputStd = sOutput.StdOut.ReadAll
			If InStr(sOutputStd, "mail exchanger") = False Then
				button = MsgBox("Error: MX record lookup failed." & vbCrLf & "Domain: " & mailDomain & vbCrLf & "Description: " & sOutput.StdErr.ReadAll, vbCritical + vbAbortRetryIgnore + vbDefaultButton2 + vbSystemModal, "DMail (VBScript)")
				Select Case button
				Case vbAbort
					WScript.Quit
				Case vbRetry
					success = False
				Case vbIgnore
					button = MsgBox("Are you sure you want to ignore this error? Ignoring may cause unexpected problems.", vbExclamation + vbYesNo + vbDefaultButton2 + vbSystemModal, "DMail (VBScript)")
					Select Case button
					Case vbYes
						success = True
					Case vbNo
						success = False
					End Select
				End Select
			Else
				success = True
			End If
		Else
			button = MsgBox("Error: MX record lookup failed." & vbCrLf & "Domain: No domain supplied." & vbCrlf & "Description: This operation can't be retried.", vbCritical + vbOKOnly + vbDefaultButton2 + vbSystemModal, "DMail (VBScript)")
			WScript.Quit
		End If
	Loop
	success = False
	SMTP = Split(Split(sOutputStd, "mail exchanger = ")(1), vbCrLf)(0)
	mail.Configuration.Fields.Item ("http://schemas.microsoft.com/cdo/configuration/smtpserver") = SMTP
	
	mail.Configuration.Fields.Item ("http://schemas.microsoft.com/cdo/configuration/smtpserverport") = 25
	
	mail.Configuration.Fields.Update
	
	Do until success = True
		mail.Send
		If Err.Number <> 0 Then
			button = MsgBox("Error: Email send failed." & vbCrLf & "To: " & recipient & vbCrLf & "From: " & from & vbCrLf & "SMTP Server: " & SMTP & vbCrLf & "Port: 25" & vbCrLf & "Description: " & Err.Description, vbCritical + vbAbortRetryIgnore + vbDefaultButton2 + vbSystemModal, "DMail (VBScript)")
			Select Case button
			Case vbAbort
				WScript.Quit
			Case vbRetry
				success = False
			Case vbIgnore
				button = MsgBox("Are you sure you want to ignore this error? Ignoring may cause unexpected problems.", vbExclamation + vbYesNo + vbDefaultButton2 + vbSystemModal, "DMail (VBScript)")
				Select Case button
				Case vbYes
					success = True
				Case vbNo
					success = False
				End Select
			End Select
		Else
			success = True
		End If
	Loop
	success = False
	
	MsgBox "Email sent successfully.", vbInformation + vbOKOnly, "DMail (VBScript)"
Else
	On Error Resume Next
	Set Email = CreateObject("CDO.Message")
	
	mail.To = recipient
	
	mail.From = from
	
	mail.Subject = subject
	
	mail.TextBody = body
	
	mail.Configuration.Fields.Item ("http://schemas.microsoft.com/cdo/configuration/sendusing") = 2
	
	mail.Configuration.Fields.Item ("http://schemas.microsoft.com/cdo/configuration/smtpserver") = customSMTP
	
	port = InputBox("Port (for custom SMTP server)", "DMail (VBScript)", "25") 'Port for custom SMTP server
	Set re = New RegExp
	With re
		.Pattern = "^[0-9]+$"
	End With
	Do Until success = True
		If re.Test(port) = True Then
			success = True
		Else
			button = MsgBox("""" & port & """ is not a valid port number.", vbCritical + vbOKOnly + vbSystemModal, "DMail (VBScript)")
			success = False
		End If
	Loop
	success = False
	mail.Configuration.Fields.Item ("http://schemas.microsoft.com/cdo/configuration/smtpserverport") = port
	
	mail.Configuration.Fields.Update
	
	Do until success = True
		mail.Send
		If Err.Number <> 0 Then
			button = MsgBox("Error: EEmail send failed." & vbCrLf & "To: " & recipient & vbCrLf & "From: " & from & vbCrLf & "SMTP Server: " & customSMTP & vbCrLf & "Port: " & port & vbCrLf & "Description: " & Err.Description & vbCrLf & "Note: If you can't send any email to this address, and both the ""To"" and ""From"" addresses are valid and existent, it is most likely the SMTP server is rejecting your email due to your IP address not being recognised by anti-spam software as an SMTP server. You can try again or alternatively get your IP address whitelisted by the recipient's SMTP server.", vbCritical + vbAbortRetryIgnore + vbDefaultButton2 + vbSystemModal, "DMail (VBScript)")
			Select Case button
			Case vbAbort
				WScript.Quit
			Case vbRetry
				success = False
			Case vbIgnore
				button = MsgBox("Are you sure you want to ignore this error? Ignoring may cause unexpected problems.", vbExclamation + vbYesNo + vbDefaultButton2 + vbSystemModal, "DMail (VBScript)")
				Select Case button
				Case vbYes
					success = True
				Case vbNo
					success = False
				End Select
			End Select
		Else
			success = True
		End If
	Loop
	success = False
	
	MsgBox "Email sent successfully.", vbInformation + vbOKOnly, "DMail (VBScript)"
End If
