Imports System.Net.Mail
Imports System.Text

Public Class DMailForm
    Private Sub RunBtn_Click(sender As Object, e As EventArgs) Handles RunBtn.Click
        On Error GoTo ErrorHandler
        Dim recipient As String = InputBox("To", "$safeprojectname$ (VB.NET)")
        Dim from As String = InputBox("From", "$safeprojectname$ (VB.NET)")
        Dim subject As String = InputBox("Subject (optional)", "$safeprojectname$ (VB.NET)")
        Dim body As String = InputBox("Body (optional)", "$safeprojectname$ (VB.NET)")
        Dim customSMTP As String = InputBox("SMTP Server (optional)", "$safeprojectname$ (VB.NET)")

        If customSMTP = Nothing Then
            Dim emaildomain As String = Split(recipient, "@")(1)

            Dim oProcess As New Process()
            Dim oStartInfo As New ProcessStartInfo("nslookup", "-q=mx " & emaildomain) With {
            .UseShellExecute = False,
            .CreateNoWindow = True,
            .RedirectStandardOutput = True,
            .RedirectStandardError = True
        }

            oProcess.StartInfo = oStartInfo
            oProcess.Start()

            Dim sOutput As String
            Using oStreamReader As System.IO.StreamReader = oProcess.StandardOutput
                sOutput = oStreamReader.ReadToEnd()
            End Using

            If sOutput = Nothing Then
                Using oStreamReader As System.IO.StreamReader = oProcess.StandardError
                    sOutput = oStreamReader.ReadToEnd()
                End Using
            End If

            Dim sOutputStripped As String = Split(sOutput, "mail exchanger = ")(1)

            Dim SMTP As New SmtpClient(Split(sOutputStripped, vbCrLf)(0))

            Dim mail As New MailMessage()

            mail.To.Add(recipient)
            mail.From = New MailAddress(from)

            mail.Subject = subject
            mail.Body = body

            SMTP.Send(mail)
            MsgBox("Email sent successfully.", vbInformation + vbOKOnly, "VBMail")
        Else
            Dim SMTP As New SmtpClient(customSMTP)

            Dim mail As New MailMessage()

            mail.To.Add(recipient)
            mail.From = New MailAddress(from)

            mail.Subject = subject
            mail.Body = body

            SMTP.Send(mail)
            MsgBox("Email sent successfully.", vbInformation + vbOKOnly, "VBMail")
        End If

ErrorHandler:
        Dim fixesFrmt As String = ""
        Dim provider As EncodingProvider = CodePagesEncodingProvider.Instance
        Encoding.RegisterProvider(provider)
        Select Case Err.Number
            Case 9
                Dim fixes() As String = {"Check the email addresses entered."}
                For i = 0 To fixes.Length - 1 Step 1
                    fixesFrmt = fixesFrmt & Chr(149) & " " & fixes(i) & vbCrLf
                    Dim count As Integer = count + 1
                    If count = fixes.Length Then
                        MsgBox("Error #" & Err.Number & ": " & Err.Description & vbCrLf & vbCrLf & "Possible Fixes:" & vbCrLf & fixesFrmt, vbCritical + vbOKOnly, "$safeprojectname$ (VB.NET)")
                    End If
                Next
            Case Else
                Dim fixes() As String = {"No fixes found."}
        End Select
        Exit Sub
    End Sub

    Private Sub GitHubBtn_Click(sender As Object, e As EventArgs) Handles GitHubBtn.Click
        Dim PSI As New ProcessStartInfo
        With PSI
            .FileName = "https://github.com/gyware/$safeprojectname$"
            .UseShellExecute = True
        End With
        Try
            Me.Cursor = Cursors.AppStarting
            Process.Start(PSI)
        Catch
        Finally
            Me.Cursor = Cursors.Default
        End Try
    End Sub

    Private Sub ExitBtn_Click(sender As Object, e As EventArgs) Handles ExitBtn.Click
        Close()
    End Sub

    Private Sub DMailForm_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Me.ActiveControl = RunBtn
    End Sub
End Class
