Option Explicit
' Usage: $1 dirname fname number

Class files_dirs
	Dim hd_app, hd_excel_out, hd_sheet_out
	Dim root_dirname, xls_out_fullpath

	Private Sub Class_Initialize()
		Set hd_app = WScript.CreateObject("Excel.Application")
		hd_app.DisplayAlerts = False
		
		hd_app.Workbooks.Add()
		Set hd_excel_out = hd_app.Workbooks(hd_app.Workbooks.Count)
		Set hd_sheet_out = hd_excel_out.Worksheets(1)
		hd_sheet_out.Activate()
		Dim ret
		ret = print_header()
	end Sub
	
	Private Function print_header()
		Dim idx_x
		idx_x = 1

		Dim this_label(30) ' as String
		this_label(0) = "File/Dir"
		this_label(1) = "SeqNo"
		this_label(2) = "Depth"
		this_label(3) = "FileType"
		this_label(4) = "Size"
		this_label(5) = "DateLastModified"
		this_label(6) = "Owner"
		this_label(7) = "Filename"
		this_label(8) = "Link_File"
		this_label(9) = "Link_Dir"
		this_label(10) = "Fullpath"
		this_label(11) = "Dirname"

		Dim idx
		for idx_x = 1 to 15
			this_label(idx_x+11) = "Path" & idx_x
		next

		for idx_x = 0 to 26
			'WScript.StdOut.WriteLine idx_x & ": " & this_label(idx_x)
			
			hd_sheet_out.Cells(1, idx_x+1) = this_label(idx_x)
		next
	end Function

	Public Function finalize_xls(this_xls_fname)
		xls_out_fullpath = this_xls_fname
		hd_excel_out.SaveAs(xls_out_fullpath)

		hd_app.Quit
		Set hd_app = Nothing
	end Function


	Public Function process(this_root_dir, this_xls_fname)
		root_dirname = this_root_dir
		xls_out_fullpath = this_xls_fname
		
		Dim root_dirname_len
		root_dirname_len = Len(root_dirname)
		Dim idx_y_output
		idx_y_output = 2
		Dim strScriptPath
		Dim line
				
		Dim cnt, seqno
		cnt = 0
		seqno = 1
		Do Until WScript.Stdin.AtEndOfStream = True or cnt >= 10000
			'line = file_hd.ReadLine
			line = WScript.StdIn.ReadLine

			Dim this_fullpath, this_filename, this_path, this_dirname, this_suffix, this_fullpath_0, this_dirname_0
			Dim this_depth, this_dirname_len, this_filename_len, idx_suffix, flg_file_dir

			if Right(line, 8) = " のディレクトリ" then
				this_dirname_0 = Trim(Mid(line, 2, Len(line) - 8))
				this_dirname = Right(this_dirname_0, Len(this_dirname_0) - root_dirname_len)
			    this_dirname_len = Len(this_dirname)
				this_path = Split(this_dirname, "\") 
			    this_depth = UBound(this_path) + 1

				
			elseif Left(line, 1) = "2" or Left(line, 1) = "1" then
				Dim re
				set re = createObject("VBScript.RegExp")
				re.pattern = "\s+"
				re.Global = True
				'Set re = New Regexp("\s+")

				Dim data, this_date, this_time, this_size, this_owner
				line = re.Replace(line, " ")
				'data = re.Split(line)
				data = Split(line)


				this_date = data(0)
				this_time = data(1)
				this_size = data(2)
				this_owner = data(3)
				this_filename = data(4)
				if UBound(data) > 4 then
					Dim idx
					for idx = 5 to UBound(data)
						this_filename = this_filename + " " + data(idx)
					next
				end if
				
			    this_filename_len = Len(this_filename)
			    
			    if Not(this_filename = "." or this_filename = "..") then
				    if this_size = "<DIR>" then
				    	flg_file_dir = "Dir"
				    else
				    	flg_file_dir = "File"
				    end if
				    
				    if this_dirname_len = 0 then
						this_dirname = "ROOT"
					    this_fullpath = this_filename
					    this_fullpath_0 = this_dirname_0 & "\" & this_filename
					else
						this_fullpath = this_dirname & "\" & this_filename
					    this_fullpath_0 = this_dirname_0 & "\" & this_filename
				    end if

					idx_suffix = inStrRev(this_filename, ".")
					if idx_suffix > 0 then
						this_suffix = Right(this_filename, Len(this_filename) - idx_suffix)
					else
						this_suffix = "-"
					end if
					
					' ###### Output
					hd_sheet_out.Cells(idx_y_output, 1) = flg_file_dir
					hd_sheet_out.Cells(idx_y_output, 2) = seqno
					hd_sheet_out.Cells(idx_y_output, 3) = this_depth+1
					hd_sheet_out.Cells(idx_y_output, 4) = this_suffix
					if flg_file_dir = "File" then
						hd_sheet_out.Cells(idx_y_output, 5) = this_size
					end if
					hd_sheet_out.Cells(idx_y_output, 6) = this_date & " " & this_time
					hd_sheet_out.Cells(idx_y_output, 7) = this_owner
					
					hd_sheet_out.Cells(idx_y_output, 8) = this_filename
					hd_sheet_out.Cells(idx_y_output, 9).Formula = "=hyperlink(" & Chr(34) & this_fullpath_0 & Chr(34) & ", " & Chr(34) & "Link" & Chr(34) & ")"
					hd_sheet_out.Cells(idx_y_output, 10).Formula = "=hyperlink(" & Chr(34) & this_dirname_0 & Chr(34) & ", " & Chr(34) & "Link" & Chr(34) & ")"
					hd_sheet_out.Cells(idx_y_output, 11) = this_fullpath
					hd_sheet_out.Cells(idx_y_output, 12) = this_dirname
					Dim idx_x
					for idx_x = 0 to 14
						if idx_x < this_depth then
							hd_sheet_out.Cells(idx_y_output, idx_x+13) = this_path(idx_x)
						elseif idx_x = this_depth then
							hd_sheet_out.Cells(idx_y_output, idx_x+13) = this_filename
						end if
					next
					
					'for idx_x = 0 to UBound(data)
					'	hd_sheet_out.Cells(idx_y_output, 30+idx_x) = data(idx_x)
					'next
					
					idx_y_output = idx_y_output + 1
					seqno = seqno + 1
					cnt = cnt + 1
				end if
			end if
		Loop

		'hd_text.Close

		'hd_sys.Quit
		'Set hd_sys = Nothing

		hd_excel_out.SaveAs(xls_out_fullpath)

		hd_app.Quit
		Set hd_app = Nothing

	end Function
end Class


Function main()
	Dim arg
	Dim root_dir
	Dim excel_fname_out
	Set arg = WScript.Arguments
	
	if arg.Count < 2 then
		WScript.StdErr.WriteLine "Usage: cscript files-dirs.vbs root_dir output_xls_name(fullpath)"
		Exit Function
	else
		root_dir = arg(0)
		excel_fname_out = arg(1)
	end if
	
	Dim hd
	Set hd = New files_dirs
	
	Dim ret
	ret = hd.process(root_dir, excel_fname_out)
	'ret = hd.finalize_xls(excel_fname_out)

end Function

main()
