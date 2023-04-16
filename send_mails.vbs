' ################################################################################################################
' #                                                                                                              #
' #                                                                                                              #
' #                                          sendmail                                                            #
' #                                                                                                              #
' #                                                                                                              #
' ################################################################################################################

Option Explicit
' Usage: type email(head & body) | cscript $1 

Function main()
	Dim arg
	Set arg = WScript.Arguments
	
	Dim email_to
	Dim email_subject
	Dim email_body
	Dim email_attach
	Dim email_flg

	' Open Outlook
	Dim handler
	Dim this_namespace
	'Dim handler As Outlook.Application
	'Dim this_namespace As Outlook.NameSpace
	Dim this_folder

	' prepare send mail
	Set handler = CreateObject("Outlook.Application")
	Set this_namespace = handler.GetNamespace("MAPI")
	this_namespace.Logon "analyticsinterestgroup",, False, True
	'this_namespace.Logon "LatestProfile",, True, True
	'this_namespace.Logon "Default Outlook Profile",, False, True
	Set this_folder = this_namespace.GetDefaultFolder(6)
	this_folder.Display
	
	handler.ActiveWindow.WindowState = 2

	Dim this_email
	Dim timestamp
	timestamp = Now
	
	' Open File & Send Email
	Dim cnt
	Dim strInput
	cnt = 0
	email_to = ""
	email_subject = ""
	email_body = ""
	email_attach = ""
	email_flg = 0
	Do Until WScript.StdIn.AtEndOfStream
		strInput = WScript.StdIn.ReadLine
		'WScript.Echo strInput	
		if Left(strInput, 1) = "#" then
			if Left(strInput, 5) = "# TO:" then
				email_to = Mid(strInput, 6, 4000)
				this_email.To = email_to
				email_flg = email_flg + 1
			elseif Left(strInput, 6) = "# SUB:" then
				email_subject = Mid(strInput, 7, 2000)
				this_email.Subject = email_subject
				'& "(" & timestamp & ")"
				email_flg = email_flg + 2
			elseif Left(strInput, 7) = "# ATCH:" then
				email_attach = Mid(strInput, 8, 2000)
				WScript.Echo email_attach
				this_email.Attachments.Add email_attach
			elseif Left(strInput, 5) = "# BEG" then
				cnt = cnt + 1
				Set this_email = handler.CreateItem(0)
				this_email.Display
			elseif Left(strInput, 5) = "# END" then
				if email_flg <> 3 then
					WScript.Echo "Error in " & cnt & "-th mail"	
				else
					' Log
					WScript.Echo "############### " & cnt & "-th mail ###################"	
					WScript.Echo "--------------- to -----------------------------"	
					WScript.Echo email_to
					WScript.Echo "---------------- subject ----------------------------"	
					WScript.Echo email_subject
					WScript.Echo "---------------- attach ----------------------------"	
					WScript.Echo email_attach
					WScript.Echo "---------------- body ----------------------------"	
					WScript.Echo email_body	
					WScript.Echo "--------------------------------------------"	
					Dim return_value
					'return_value = send_mail(handler, email_to, email_subject, email_body)
					
					' Send Email
					this_email.Body = email_body
					this_email.Save
					this_email.Send
					
					email_to = ""
					email_subject = ""
					email_body = ""
					email_attach = ""
					email_flg = 0
					WScript.Echo "---------------- Waiting 2 seconds ----------------------------"	
					WScript.Sleep 1000
				end if
			end if
		else
			email_body = email_body & strInput & vbCrLf
		end if
	Loop

	WScript.Echo "---------------- Waiting 30 seconds ----------------------------"	
	WScript.Sleep 30000
	handler.Quit

end Function


Function temp_mail()	
	Set handler = CreateObject("Outlook.Application")
	Set this_namespace = handler.GetNamespace("MAPI")
	this_namespace.Logon "analyticsinterestgroup",, False, True
	'this_namespace.Logon "LatestProfile",, True, True
	'this_namespace.Logon "Default Outlook Profile",, False, True
	Set this_folder = this_namespace.GetDefaultFolder(6)
	this_folder.Display
	
	handler.ActiveWindow.WindowState = 2
end Function

Function temp_mail2()	
	Set handler = CreateObject("Outlook.Application")

	handler.Quit
end Function

Function send_mail(handler, email_to, email_subject, email_body)	
	Dim this_email
	Dim attach_fname
	Dim timestamp
	
	timestamp = Now
	attach_fname = "C:\Users\kfehvb1\Documents\Docs\Temp\test.xlsx"
	
	Set this_email = handler.CreateItem(0)
	this_email.Display

	this_email.Subject = email_subject & "(" & timestamp & ")"
	this_email.To = email_to
	this_email.Body = email_body
	this_email.Attachments.Add attach_fname
	
	this_email.Save
	this_email.Send


	send_mail = 0
end Function



main()

