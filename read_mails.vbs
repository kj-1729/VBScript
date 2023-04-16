' ################################################################################################################
' #                                                                                                              #
' #                                                                                                              #
' #                                          readmail                                                            #
' #                                                                                                              #
' #                                                                                                              #
' ################################################################################################################

Option Explicit
' Usage: type cscript $1 
'On Error Resume Next

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
    Dim mITEM
	
	' prepare send mail
	Set handler = CreateObject("Outlook.Application")
	Set this_namespace = handler.GetNamespace("MAPI")
	this_namespace.Logon "analyticsinterestgroup",, False, True
	'this_namespace.Logon "LatestProfile",, True, True
	'this_namespace.Logon "Default Outlook Profile",, False, True
	Set this_folder = this_namespace.GetDefaultFolder(6)
	this_folder.Display
	Dim n
	
	handler.ActiveWindow.WindowState = 2

    'For n = 1 To 10 'アイテム数分ループ
	Dim max_count
	max_count = this_folder.Items.Count
	WScript.StdErr.WriteLine "Main	" & max_count
	
	' Header
	WScript.Echo  "SenderName" & "	" & "SenderEmailAddress" & "	" & "FolderName" & "	" & "SeqNo" & "	" & "ReceivedTime" & "	" & "Subject"

	' Mail Items
    'For n = 1 To 10 'アイテム数分ループ
    For n = 1 To this_folder.Items.Count 'アイテム数分ループ
		'WScript.StdErr.WriteLine n
        Set mITEM = this_folder.Items(n)
        '↑代入が終わったので、各プロパティに mITEM.XXXX で アクセスする
        'Debug.Print "件名:" & mITEM.Subject  '件名表示
		
		if TypeName(mITEM) = "MailItem" then
			WScript.Echo  mITEM.SenderName & "	" & mITEM.SenderEmailAddress & "	" & "Main" & "	" & n & "	" & mITEM.ReceivedTime & "	" & mITEM.Subject
		else
			WScript.StdErr.Write "Main" & "	" & n & "	" & VarType(mITEM) & "	" & TypeName(mITEM) & "	"
		
		'if Err.Number <> 0 then
		'	WScript.StdErr.WriteLine "Error1 in " & n
		'	WScript.StdErr.WriteLine  "Main" & "	" & mITEM.SenderName & "	" & mITEM.SenderEmailAddress & "	" & mITEM.Subject
		'else
		'	WScript.StdErr.WriteLine "Error2 in " & n
		'	WScript.StdErr.WriteLine  "Main" & "	" & mITEM.SenderName & "	" & mITEM.SenderEmailAddress & "	" & mITEM.Subject
		end if
    Next
	

	' Sub Folders
	Dim c
	Dim subFolder
    For c = 1 To this_folder.Folders.Count  'サブフォルダーの数だけループする
        Set subFolder = this_folder.Folders.Item(c) 'c番目のフォルダーを代入
        Set handler.ActiveExplorer.CurrentFolder = subFolder  '移動
        
        'サブフォルダーのメール数分ループ
        'Debug.Print "サブフォルダ名: " & subFolder.Name & " には、"
        'Debug.Print "メールが " & subFolder.Items.Count & "通"
        
		max_count = subFolder.Items.Count
		WScript.StdErr.WriteLine subFolder.Name & "	" & max_count
        'For n = 1 To 10 'アイテム数分ループ
        For n = 1 To subFolder.Items.Count 'アイテム数分ループ
            Set mITEM = subFolder.Items(n)
			if TypeName(mITEM) = "MailItem" then
				WScript.Echo  mITEM.SenderName & "	" & mITEM.SenderEmailAddress & "	" & subFolder.Name & "	" & n & "	" & mITEM.ReceivedTime & "	" & mITEM.Subject
			else
				WScript.StdErr.Write subFolder.Name & "	" & n & "	" & VarType(mITEM) & "	" & TypeName(mITEM) & "	"
			end if
        Next
    Next
	
	
	'WScript.Echo "---------------- Waiting 30 seconds ----------------------------"	
	WScript.Sleep 10000
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

