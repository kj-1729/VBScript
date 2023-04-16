Option Explicit
' Usage: $1 dirname fname number

Class read_excel
	Dim hd_app
	Dim hd_excel_in, hd_sheet_in
	Dim hd_excel_out, hd_sheet_out
	Dim outdelim


	Private Sub Class_Initialize()
		Set hd_app = WScript.CreateObject("Excel.Application")
		hd_app.DisplayAlerts = False
		outdelim = Chr(9)
		
		hd_app.Workbooks.Add()
		Set hd_excel_out = hd_app.Workbooks(hd_app.Workbooks.Count)
		Set hd_sheet_out = hd_excel_out.Worksheets(1)
		hd_sheet_out.Activate()
		ret = print_header()
	end Sub

	Public Function finalize(xls_out_fullpath)
		hd_excel_out.SaveAs(xls_out_fullpath)

		hd_app.Quit
		Set hd_app = Nothing
	end Function
	
	Private Function print_header()
		Dim idx_x
		idx_x = 1
		
		hd_sheet_out.Cells(1, idx_x) ="filename"
		idx_x = idx_x + 1
		hd_sheet_out.Cells(1, idx_x) ="sheet_idx"
		idx_x = idx_x + 1
		hd_sheet_out.Cells(1, idx_x) ="sheet_name"
		idx_x = idx_x + 1
		hd_sheet_out.Cells(1, idx_x) ="row_idx"
		idx_x = idx_x + 1
		hd_sheet_out.Cells(1, idx_x) ="col_idx"
		idx_x = idx_x + 1
		hd_sheet_out.Cells(1, idx_x) ="format"
		idx_x = idx_x + 1
		hd_sheet_out.Cells(1, idx_x) = "value"
	end Function

	Public Function process(xls_in_fullpath)
		Dim excel_fname, dirname, this_idx
		
		this_idx = inStrRev(xls_in_fullpath, "\")
		if this_idx > 0 then
			excel_fname = Mid(xls_in_fullpath, this_idx+1)
			dirname = Left(xls_in_fullpath, this_idx-1)
		else
			excel_fname = xls_in_fullpath
			dirname = ""
		end if
				
		WScript.StdErr.WriteLine "fullpath: " & xls_in_fullpath
		WScript.StdErr.WriteLine "dirname: " & dirname
		WScript.StdErr.WriteLine "filenamee: " & excel_fname
		
		' (VBA) : https://docs.microsoft.com/ja?jp/office/vba/api/excel.workbooks.
		' updatelinks=0 ( )
		Set hd_excel_in =hd_app.WorkBooks.Open(xls_in_fullpath, 0)
		
		' ####### Temporary
		Dim sheet_idx, num_sheets
		num_sheets = hd_excel_in.Worksheets.Count
		
		Dim ret
		Dim idx_y_output
		idx_y_output = 2
		for sheet_idx = 1 to num_sheets
			idx_y_output = process_sheet(excel_fname, sheet_idx, idx_y_output)
		next

		hd_excel_in.Close False
	end Function

	Public Function process_sheet(excel_fname, sheet_idx, idx_y_output)
		WScript.StdErr.WriteLine "Process Sheet " & sheet_idx
		
		Set hd_sheet_in = hd_excel_in.WorkSheets.Item(sheet_idx)
		Dim cell_range
		cell_range = hd_sheet_in.UsedRange
		
		if isempty(cell_range) then
			WScript.StdErr.WriteLine sheet_idx & ": " & "Empty"
			Exit Function
		end if
		
		Dim idx_row, idx_col
		for idx_row = 1 to UBound(cell_range, 1)
			for idx_col = 1 to UBound(cell_range, 2)
				Dim this_cell, this_format, this_formula, this_value
				Set this_cell = hd_sheet_in.Cells(idx_row, idx_col)
				' ############## Temporary ###################
				On Error Resume Next
				this_formula = this_cell.Formula
				' #### Temporary
				if Len(this_formula) > 0 and Left(this_formula, 1) <> "=" then
					this_format = this_cell.NumberFormatLocal
					this_value = Replace(this_cell.Value, VbCrLf, "<RET>", 1, -1, 1)
					this_value = Replace(this_value, VbCr, "<RET>", 1, -1, 1)
					this_value = Replace(this_value, VbLf, "<RET>", 1, -1, 1)
					if Len(this_value) > 0 then
						hd_sheet_out.Cells(idx_y_output, 1) = excel_fname
						hd_sheet_out.Cells(idx_y_output, 2) = sheet_idx
						hd_sheet_out.Cells(idx_y_output, 3) = hd_sheet_in.name
						hd_sheet_out.Cells(idx_y_output, 4) = idx_row
						hd_sheet_out.Cells(idx_y_output, 5) = idx_col
						hd_sheet_out.Cells(idx_y_output, 6) = this_format
						hd_sheet_out.Cells(idx_y_output, 7) = this_value
						idx_y_output = idx_y_output + 1
					end if
				end if
				On Error GoTo 0
				' ############## Temporary ###################
			next
		next
		process_sheet = idx_y_output
	end Function
end Class

Function tmp()
	Dim excel_fname
	Dim obj, excel, sheet
	excel_fname = "E:YusersYpapaYDocsYTempYtest.xlsx"
	Set obj = WScript.CreateObject("Excel.Application")
	Set excel = obj.WorkBooks.Open(excel_fname)
	Set sheet = excel.WorkSheets.Item(1)
	WScript.Echo sheet.Cells(1, 1)
	obj.Quit()
end Function

Function main()
	Dim arg
	Dim dirname
	Dim excel_fname_in, excel_fname_out
	Set arg = WScript.Arguments
	if arg.Count < 2 then
		WScript.StdErr.WriteLine "Usage: cscript read_excel.vbs excel_fname(fullpath) output_xls_name(fullpath)"
		Exit Function
	else
		excel_fname_in = arg(0)
		excel_fname_out = arg(1)
	end if
	
	Dim hd
	Set hd = New read_excel
	Dim ret
	ret = hd.process(excel_fname_in)
	ret = hd.finalize(excel_fname_out)

end Function

main()
