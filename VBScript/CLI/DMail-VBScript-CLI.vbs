 Set oShell = CreateObject("WScript.Shell")
 vbsInterpreter = "cscript.exe"

 Call ForceConsole()

 Function printf(txt)
    WScript.StdOut.WriteLine txt
 End Function

 Function printl(txt)
    WScript.StdOut.Write txt
 End Function

 Function scanf()
    scanf = LCase(WScript.StdIn.ReadLine)
 End Function

 Function wait(n)
    WScript.Sleep Int(n * 1000)
 End Function

 Function exitc()
    WScript.StdOut.Write "Press enter to exit..."
	WScript.StdIn.ReadLine
	WScript.Quit
 End Function

 Function ForceConsole()
    If InStr(LCase(WScript.FullName), vbsInterpreter) = 0 Then
        oShell.Run vbsInterpreter & " //NoLogo " & Chr(34) & WScript.ScriptFullName & Chr(34)
        WScript.Quit
    End If
 End Function

 Function cls()
    For i = 1 To 50
        printf ""
    Next
 End Function

If WScript.Arguments.Item(0) = "about" Then
printf "DMail VBScript CLI [https://github.com/gyware/dmail]"
printf "Copyright (C) GyWare. All rights reserved."
Else


Set nameArgs = WScript.Arguments.Named

recipients = nameArgs.Item("to") 'To

from = nameArgs.Item("from") 'From

subject = nameArgs.Item("subject") 'Subject (optional)

body = nameArgs.Item("body") 'Body (optional)

customSMTP = nameArgs.Item("smtp") 'Custom SMTP Server (optional)


Set mail = CreateObject("CDO.Message")

mail.To = recipients

mail.From = from

mail.Subject = subject

mail.TextBody = body

If customSMTP = "" Then
	mail.Configuration.Fields.Item ("http://schemas.microsoft.com/cdo/configuration/sendusing") = 2
	
	On Error Resume Next
	
		mailDomain = Split(Mail.To, "@")(1)
		If Err.Number <> 0 Then
			printf "Error: Domain extraction failed." & vbCrLf & "Description: " & Err.Description & vbCrLf & "Note: The ""To"" address may be invalid and is hence causing this error."
			WScript.Quit
		End If
	
	Set oShell = CreateObject("WScript.Shell")
		If mailDomain <> "" Then
			Set sOutput = oShell.Exec("nslookup -q=mx " & mailDomain)
			sOutputStd = sOutput.StdOut.ReadAll
			If InStr(sOutputStd, "mail exchanger") = False Then
				printf "Error: MX record lookup failed." & vbCrLf & "Domain: " & mailDomain & vbCrLf & "Description: " & sOutput.StdErr.ReadAll
				WScript.Quit
			End If
		Else
			printf "Error: MX record lookup failed." & vbCrLf & "Domain: No domain supplied."
			WScript.Quit
		End If
	SMTP = Split(Split(sOutputStd, "mail exchanger = ")(1), vbCrLf)(0)
	mail.Configuration.Fields.Item ("http://schemas.microsoft.com/cdo/configuration/smtpserver") = SMTP
	
	mail.Configuration.Fields.Item ("http://schemas.microsoft.com/cdo/configuration/smtpserverport") = 25
	
	mail.Configuration.Fields.Update
	
		mail.Send
		If Err.Number <> 0 Then
			printf "Error: Email send failed." & vbCrLf & "SMTP Server: " & SMTP & vbCrLf & "Port: 25" & vbCrLf & "Description: " & Err.Description
			WScript.Quit
		End If
	
	printf "Email sent successfully to """ & recipients & """."
Else
	On Error Resume Next
	Set Email = CreateObject("CDO.Message")
	
	mail.To = recipients
	
	mail.From = from
	
	mail.Subject = subject
	
	mail.TextBody = body
	
	mail.Configuration.Fields.Item ("http://schemas.microsoft.com/cdo/configuration/sendusing") = 2
	
	mail.Configuration.Fields.Item ("http://schemas.microsoft.com/cdo/configuration/smtpserver") = customSMTP
	
	printl "Port (for custom SMTP server)" 'Port for custom SMTP server
	port = scanf
	mail.Configuration.Fields.Item ("http://schemas.microsoft.com/cdo/configuration/smtpserverport") = port
	
	mail.Configuration.Fields.Update
	
		mail.Send
		If Err.Number <> 0 Then
			printf "Error: Email send failed." & vbCrLf & "SMTP Server: " & customSMTP & vbCrLf & "Port: " & port & vbCrLf & "Description: " & Err.Description & vbCrLf & "Note: If you can't send any email to this address, and both the ""To"" and ""From"" addresses are valid and existent, it is most likely the SMTP server is rejecting your email due to your IP address not being recognised by anti-spam software as an SMTP server. You can try again or alternatively get your IP address whitelisted by the recipient's SMTP server."
			WScript.Quit
		End If
	
	printf "Email sent successfully to """ & recipients & """."
End If
End If