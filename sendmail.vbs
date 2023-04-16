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

	Dim cnt
	Dim strInput
	cnt = 0
	email_to = ""
	email_subject = ""
	email_body = ""
	Do Until WScript.StdIn.AtEndOfStream
		strInput = WScript.StdIn.ReadLine
		if cnt = 0 then
			email_to = strInput
		elseif cnt = 1 then
			email_subject = strInput
		else
			email_body = email_body & strInput & vbCrLf
		end if
		cnt = cnt + 1
	Loop

	WScript.Echo "--------------- to -----------------------------"	
	WScript.Echo email_to
	WScript.Echo "---------------- subject ----------------------------"	
	WScript.Echo email_subject
	WScript.Echo "---------------- body ----------------------------"	
	WScript.Echo email_body	
	WScript.Echo "--------------------------------------------"	

	Dim return_value
	return_value = send_mail(email_to, email_subject, email_body)

end Function


Function send_mail(email_to, email_subject, email_body)
	Dim handler
	Dim this_namespace
	Dim this_folder
	Dim this_email

	' prepare send mail
	Set handler = CreateObject("Outlook.Application")

	Set this_namespace = handler.GetNamespace("MAPI")
	Set this_folder = this_namespace.GetDefaultFolder(6)

	this_folder.Display
	handler.ActiveWindow.WindowState = 2

	Set this_email = handler.CreateItem(0)
	this_email.Display

	this_email.Subject = email_subject
	this_email.To = email_to
	this_email.Body = email_body
	'this_email.Attachments.Add attach_fname
	
	this_email.Save
	this_email.Send

	handler.Quit

	send_mail = 0
end Function

main()

