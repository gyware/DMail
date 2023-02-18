 Set oWSH = CreateObject("WScript.Shell")
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
        oWSH.Run vbsInterpreter & " //NoLogo " & Chr(34) & WScript.ScriptFullName & Chr(34)
        WScript.Quit
    End If
 End Function

 Function cls()
    For i = 1 To 50
        printf ""
    Next
 End Function

printf "DMail (VBScript) CLI [https://github.com/gyware/dmail]"
printf "Copyright (C) GyWare. All rights reserved."
printf ""

Set mail = CreateObject("CDO.Message")

printl "To: " 'To
mail.To = scanf

printl "From: " 'From
mail.From = scanf

printl "Subject (optional): " 'Subject
mail.Subject = scanf

printl "Body (optional): " 'Body
mail.TextBody = scanf

printl "SMTP Server (optional): " 'Custom SMTP Server
customSMTP = scanf

If customSMTP = "" Then
	mail.Configuration.Fields.Item ("http://schemas.microsoft.com/cdo/configuration/sendusing") = 2
	
	On Error Resume Next
	
		mailDomain = Split(Mail.To, "@")(1)
		If Err.Number <> 0 Then
			printf "Error: Domain extraction failed." & vbCrLf & "Description: " & Err.Description & vbCrLf & "Note: The ""To"" address may be invalid hence causing this error."
			exitc
		End If
	
	Set oShell = CreateObject("WScript.Shell")
		If mailDomain <> "" Then
			Set sOutput = oShell.Exec("nslookup -q=mx " & mailDomain)
			sOutputStd = sOutput.StdOut.ReadAll
			If InStr(sOutputStd, "mail exchanger") = False Then
				printf "Error: MX record lookup failed." & vbCrLf & "Domain: " & mailDomain & vbCrLf & "Description: " & sOutput.StdErr.ReadAll
				exitc
			End If
		Else
			printf "Error: MX record lookup failed." & vbCrLf & "Domain: No domain supplied."
			exitc
		End If
	SMTP = Split(Split(sOutputStd, "mail exchanger = ")(1), vbCrLf)(0)
	mail.Configuration.Fields.Item ("http://schemas.microsoft.com/cdo/configuration/smtpserver") = SMTP
	
	mail.Configuration.Fields.Item ("http://schemas.microsoft.com/cdo/configuration/smtpserverport") = 25
	
	mail.Configuration.Fields.Update
	
		mail.Send
		If Err.Number <> 0 Then
			printf "Error: Email send failed." & vbCrLf & "SMTP Server: " & SMTP & vbCrLf & "Port: 25" & vbCrLf & "Description: " & Err.Description
			exitc
		End If
	
	printf "Email sent successfully."
	exitc
Else
	On Error Resume Next
	Set Email = CreateObject("CDO.Message")
	
	mail.To = recipient
	
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
			exitc
		End If
	
	printf "Email sent successfully."
	exitc
End If
