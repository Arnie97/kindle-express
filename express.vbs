Title    = "Kindle Express"

Sender   = "Arnie19970811@126.com"
Receiver = "Illyasviel@kindle.cn"

User     = Left(Sender, InStr(Sender, "@") - 1)
Domain   = Mid (Sender, InStr(Sender, "@") + 1)
Password = Right(User, 8)

Count = WScript.Arguments.Count
If Count = 0 Then
	MsgBox "No input files.", vbCritical, Title
	WScript.Quit
End If

With CreateObject("CDO.Message")
	With .Configuration.Fields
		NameSpace = "http://schemas.microsoft.com/cdo/configuration/"
		.Item(NameSpace & "sendusername")     = User
		.Item(NameSpace & "sendpassword")     = Password
		.Item(NameSpace & "sendusing")        = 2
		.Item(NameSpace & "smtpauthenticate") = 1
		.Item(NameSpace & "smtpserver")       = "smtp." + Domain
		.Item(NameSpace & "smtpserverport")   = 25
		.Update
	End With

	Do
		Receiver = InputBox("To:", Title, Receiver)
		If Receiver = vbNullString Then
			If MsgBox("Cancel operation?", vbYesNo + vbExclamation, Title) = vbYes Then
				WScript.Quit
			End If
		Else
			Exit Do
		End If
	Loop
	.To = Receiver

	List = vbNullString
	For Each File in WScript.Arguments
		.AddAttachment File
		List = List & File & vbCrLf
	Next

	.From = Sender
	.Subject = Title
	.TextBody = Title

	On Error Resume Next
	Do
		.Send
		If Err.Number = 0 Then
			Exit Do
		ElseIf MsgBox("Operation failed: " & Err.Description, vbRetryCancel + vbCritical, Title) = vbCancel Then
			WScript.Quit
		End If
	Loop
End With

MsgBox List & vbCrLf & Count & " file(s) sent to " & Receiver, vbInformation, Title
