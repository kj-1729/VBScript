' ################################################################################################################
' #                                                                                                              #
' #                                                                                                              #
' #                                        Access WWW using IE                                                   #
' #                                                                                                              #
' #                                                                                                              #
' ################################################################################################################

' reference: http://qiita.com/nezuq/items/93390c4def01991c0354
' save pdf: http://foundknownanddone.dots2pattern.com/2014/05/download-pdf-file-from-ie-by-vbscript-whs.html

Option Explicit
' Usage: $1 url

Function main()
	Dim arg
	Dim url
	
	Set arg = WScript.Arguments
	if arg.Count < 1 then
		WScript.Echo "Usage:  ie.vbs url"
		Exit Function
	else
		url = arg(0)
	end if

	' Open IE
	Dim ie
	'ie = open_ie()
	Set ie = CreateObject("InternetExplorer.Application")
	ie.Visible = True

	Dim cnt
	cnt = 0
	Dim elm_list
	Dim label
	if 1 = 1 then
		WScript.StdOut.WriteLine "===" & cnt & "==="
		WScript.StdOut.WriteLine url
		WScript.StdOut.WriteLine "=================="

		if cnt >= 0 then
			Dim ret
			ret = access_url(ie, url)
		
			Dim wait_seconds
			wait_seconds = 3
			Wscript.sleep(wait_seconds*1000)
		
			Dim strHTML
			'strHTML = get_html(ie)
			'strHTML = get_body(ie)
			strHTML = ie.Document.all.tags("HTML")(0).innerHTML
		
			' Store file
			Dim fname
			Dim cnt_str
			cnt_str = Replace(Space(3 - Len(cnt)) & cnt, Space(1), "0")
			fname = "page_" & cnt_str & ".txt"
			Call save_to_file(fname, strHTML)
		end if
		
		cnt = cnt + 1
		
		Set elm_list = ie.Document.getElementsByTagName("a")
		Dim idx
		Dim this_loop
		for idx = 0 to elm_list.Length - 1
			label = pretty_string(elm_list(idx).textContent)
			
			if idx = 35 then
				WScript.StdOut.WriteLine idx & "Skip" & Len(label)
				WScript.StdOut.WriteLine elm_list(idx).textContent
				for this_loop = 1 to Len(label)
					WScript.Stdout.Write "|"
					WScript.Stdout.Write Chr(Asc(Mid(label, this_loop, 1))) 
					WScript.Stdout.Write "(" & Asc(Mid(label, this_loop, 1)) & ")"
				next
			else
				WScript.StdErr.Write idx
				WScript.StdOut.Write idx
				WScript.StdOut.Write VbTab 
				WScript.StdOut.Write elm_list(idx).href 
				WScript.StdOut.Write VbTab
				WScript.StdOut.Write Len(elm_list(idx).textContent)
				
				'WScript.StdOut.Write VbTab
				'WScript.StdOut.Write elm_list(idx).textContent
				WScript.StdOut.Write VbTab
				WScript.StdOut.Write label
				WScript.StdOut.WriteLine ""
			end if
			
			'for this_loop = 1 to Len(label3)
			'	WScript.Stdout.Write "|" & Mid(label3, this_loop, 1) & "(" & Asc(Mid(label3, this_loop, 1)) & ")"
			'next
			'WScript.StdOut.WriteLine "-----------------"
			'elm_list(idx).Click
			'Wscript.sleep(wait_seconds*1000)
		next
		Wscript.sleep(wait_seconds*1000)
		
	end if
		
	' Close ie
	Call close_ie(ie)
end Function

' ################################################################
' #                                                              #
' #               Help Functions                                 #
' #                                                              #
' ################################################################

Function pretty_string(str)
	str = Trim(str)
	str = Replace(str, VbTab, "")
	str = Replace(str, vbCrLf, " ")
	str = Replace(str, vbCr, " ")
	str = Replace(str, vbLf, " ")
	
	pretty_string = str
end Function

' ################################################################
' #                                                              #
' #               Functions to manage HTML                       #
' #                                                              #
' ################################################################
Function get_html(ie)
	Dim strHTML
	strHTML = ie.Document.all.tags("HTML")(0).innerHTML
end Function

Function get_body(ie)
	Dim strBody
	strBody = ie.Document.Body.InnerHtml
end Function

Function get_element_by_id(ie, this_id)
	Dim elm
	Set elm = ie.document.getElementById(this_id)
	'Set elm = ie.document.getElementsById(this_id)
end Function

Function get_element_by_name(ie, this_name)
	Dim elm
	Set elm = ie.document.getElementByName(this_name)
	'Set elm = ie.document.getElementsByName(this_name)
end Function

Function get_element_by_classname(ie, this_name)
	Dim elm
	Set elm = ie.document.getElementByClassName(this_name)
	'Set elm = ie.document.getElementsByClassName(this_name)
end Function

Function get_element_by_tagname(ie, this_name)
	Dim elm
	Set elm = ie.document.getElementByTagName(this_name)
	'Set elm = ie.document.getElementsByTagName(this_name)
end Function

Function focus_element(elm)
	elm.Focus
end Function

Function set_value(elm, this_value)
	elm.Value = this_value
end Function

Function elm_select(elm, this_idx)
	elm.selectedIndex = this_idx
end Function

Function elm_click(elm)
	elm.Click
end Function


' ################################################################
' #                                                              #
' #             File I/O                                         #
' #                                                              #
' ################################################################
Function save_to_file(fname, this_object)
	Dim fso, tf
	Set fso = CreateObject("Scripting.FileSystemObject")
	Set tf = fso.CreateTextFile(fname, 2, True)

	tf.Write this_object
	tf.Close
end Function


' ################################################################
' #                                                              #
' #             Browser (IE)                                     #
' #                                                              #
' ################################################################

Function open_ie()
	Dim ie, ret_ie
	Set ie = CreateObject("InternetExplorer.Application")
	ie.Visible = True
	
	ret_ie = ie
end Function


Function access_url(ie, url)
	ie.Navigate url
	Do While ie.Busy = True or ie.readyState <> 4
	Loop
end Function

Function close_ie(ie)
	ie.Quit
	Set ie = Nothing
end Function


Function full_screen(ie)
	ie.FullScreen = True
end Function

' ################################################################
' #                                                              #
' #                                                              #
' #                                                              #
' ################################################################


main()
