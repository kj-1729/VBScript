Option Explicit
' Usage: $1 dirname fname number

' ####################################################################################
' #                                                                                  #
' #                         naivebayes                                               #
' #                                                                                  #
' ####################################################################################
Class naivebayes
	Dim freq_by_word
	Dim docs_in_cat, words_in_cat
	Dim vocabularies
	Dim num_vocs, num_cats
	Dim indelim, outdelim
	
	Private Sub Class_Initialize()
		Set freq_by_word = WScript.CreateObject("Scripting.Dictionary")
		Set docs_in_cat = WScript.CreateObject("Scripting.Dictionary")
		Set words_in_cat = WScript.CreateObject("Scripting.Dictionary")
		Set vocabularies = WScript.CreateObject("Scripting.Dictionary")
		'Set vocabularies = CreateObject("System.Collections.ArrayList")

		num_vocs = 0
		num_cats = 0
		indelim = Chr(9)
		outdelim = Chr(9)
	end Sub

	' ########################################################
	' #                                                      #
	' #                 train                                #
	' #                                                      #
	' ########################################################
	Public Function train(fullpath_train)
		Dim file_hd_0
		Dim file_hd
		Set file_hd_0 = CreateObject("Scripting.FileSystemObject")
		Set file_hd = file_hd_0.OpenTextFile(fullpath_train, 1)

		Dim cnt, line, num_docs
		Dim data
		
		Dim prev_id
		cnt = 0
		num_docs = 0
		prev_id = ""

		' Header
		line = file_hd.ReadLine
		'line = WScript.StdIn.ReadLine

		' ### Input
		' 0: Document ID
		' 1: Class
		' 2: Word
		Do Until file_hd.AtEndOfStream = True
		'Do Until WScript.StdIn.AtEndOfStream
			line = file_hd.ReadLine
			'line = WScript.StdIn.ReadLine
			data =Split(line, indelim)

			if prev_id <> data(0) then
				' docs_in_cat
				if docs_in_cat.Exists(data(1)) then
					docs_in_cat.Item(data(1)) = docs_in_cat.Item(data(1)) + 1
				else
					docs_in_cat.Add data(1), 1
				end if
				prev_id = data(0)
				num_docs = num_docs + 1
			end if
			
			' count words in cat
			if words_in_cat.Exists(data(1)) then
				words_in_cat.Item(data(1)) = words_in_cat.Item(data(1)) + 1
			else
				words_in_cat.Add data(1), 1
				num_cats = num_cats + 1
			end if
			
			Dim this_category
			if freq_by_word.Exists(data(1)) then
				Set this_category = freq_by_word.Item(data(1))
			else
				Set this_category = WScript.CreateObject("Scripting.Dictionary")
				freq_by_word.Add data(1), this_category
			end if
			
			if this_category.Exists(data(2)) then
				this_category.Item(data(2)) = this_category.Item(data(2)) + 1
			else
				this_category.Add data(2), 1
			end if

			if not vocabularies.Exists(data(2)) then
				num_vocs = num_vocs + 1
				vocabularies.Add data(2), 1
			end if

			Set this_category = Nothing
			
			cnt = cnt + 1
		Loop
		
		Dim this_cat
		for each this_cat in docs_in_cat
			docs_in_cat.Item(this_cat) = docs_in_cat.Item(this_cat) / num_docs
		next

		file_hd.Close
		Set file_hd = Nothing
		Set file_hd_0 = Nothing

	end function	

	' ########################################################
	' #                                                      #
	' #                   test                               #
	' #                                                      #
	' ########################################################
	Public Function test()
		Dim words
		Set words = WScript.CreateObject("Scripting.Dictionary")
		'Set words = CreateObject("System.Collections.ArrayList")
		
		Dim this_class, posterior
		Dim cnt, line
		Dim data		
		Dim prev_id
		Dim prev_class
		cnt = 0
		prev_id = ""
		prev_class = ""
		
		WScript.StdOut.WriteLine "ID" & outdelim & "Class" & outdelim & "EstimatedClass" & outdelim & "Probability"

		' Header
		line = WScript.StdIn.ReadLine

		' ### Input
		' 0: Document ID
		' 1: Class(Dummy)
		' 2: Word
		Do Until WScript.StdIn.AtEndOfStream
			line = WScript.StdIn.ReadLine
			data =Split(line, indelim)

			if prev_id <> data(0) then
				if cnt > 0 then
					Set posterior = get_posterior(words.Keys())
					
					'Dim keywords()
					'Dim values()
					'ReDim keywords(posterior.Count-1)
					'ReDim values(posterior.Count-1)
					
					'Dim max_class, idx_class
					'idx_class = 0
					'for each this_class in posterior
					'	keywords(idx_class) = this_class
					'	values(idx_class) = posterior.Item(this_class)
					'next
					
					'call posterior.Keys.CopyTo(keywords, 0)
					'call posterior.Values.CopyTo(values, 0)
					'call Array.Sort(values, keywords)
					'call Array.Reverse(values, keywords)
					
					'if posterior.Count >= 3 then
					'	max_class = 2
					'else
					'	max_class = posterior.Count-1
					'end if

					' ### output
					'WScript.StdOut.WriteLine prev_id & outdelim & prev_class
					'for idx_class = 0 to max_class
					'	WScript.StdOut.Write outdelim & keywords(idx_class) & outdelim & values(idx_class)
					'next

					for each this_class in posterior
						WScript.StdOut.WriteLine prev_id & outdelim & prev_class & outdelim & this_class & outdelim & posterior.Item(this_class)
					next
					Set posterior = Nothing
					
					'this_class = classify(words)
					'WScript.StdOut.WriteLine prev_id & outdelim & this_class
					
					Set words = Nothing
					Set words = WScript.CreateObject("Scripting.Dictionary")
					'Set words = CreateObject("System.Collections.ArrayList")
				end if
				prev_id = data(0)
				prev_class = data(1)
			end if

			if not words.Exists(data(2)) then
				words.Add data(2), 1
			end if
			'words.Add data(2)
			
			cnt = cnt + 1
		Loop

		Set posterior = get_posterior(words.Keys())
		for each this_class in posterior
			WScript.StdOut.WriteLine prev_id & outdelim & prev_class & outdelim & this_class & outdelim & posterior.Item(this_class)
		next

		Set posterior = Nothing
		Set words = Nothing

	end function

	' ########################################################
	' #                                                      #
	' #                 print_trained_data                   #
	' #                                                      #
	' ########################################################
	Public Function print_trained_data()
		Dim this_cat, this_word, this_category
		
		WScript.StdOut.WriteLine "### VOCABULARIES: " & num_vocs & " ###"
		WScript.StdOut.WriteLine "Word"
		for each this_word in vocabularies
			WScript.StdOut.WriteLine this_word
		next
		WScript.StdOut.WriteLine ""

		WScript.StdOut.WriteLine "### CATCOUNT: " & num_cats & " ###"
		WScript.StdOut.WriteLine "Category" & outdelim & "#Documents" & outdelim & "#Words"
		for each this_cat in docs_in_cat
			WScript.StdOut.WriteLine this_cat & outdelim & docs_in_cat.Item(this_cat) & outdelim & words_in_cat.Item(this_cat)
		next
		WScript.StdOut.WriteLine ""

		WScript.StdOut.WriteLine "### freq_by_word ###"
		WScript.StdOut.WriteLine "Category" & outdelim & "Word" & outdelim & "Frequency"
		for each this_cat in freq_by_word
			Set this_category = freq_by_word.Item(this_cat)
			for each this_word in this_category
				WScript.StdOut.WriteLine this_cat & outdelim & this_word & outdelim & this_category.Item(this_word)
			next
			Set this_category = Nothing
		next
	end Function

	' ########################################################
	' #                                                      #
	' #                 read_trained_data                    #
	' #                                                      #
	' ########################################################
	Public Function read_trained_data(fname_trained_data)
		' ### vocabularies
		' ### docs_in_cat, words_in_cat
		' ### freq_by_word
	end Function
	
	' ########################################################
	' #                                                      #
	' #               classify                               #
	' #                                                      #
	' ########################################################
	Public Function classify(words)
		Dim max_score, this_score
		Dim this_cat
		
		max_score = -1000000
		
		for each this_cat in docs_in_cat
			this_score = get_score(words, this_cat)
			if this_score > max_score then
				max_score = this_score
				classify = this_cat
			end if
		next
	end function
	
	' ########################################################
	' #                                                      #
	' #             get_posterior                            #
	' #                                                      #
	' ########################################################
	Public function get_posterior(words)
		Dim sumprob, this_prob, scr, padding, cnt, this_cat
		Set get_posterior = WScript.CreateObject("Scripting.Dictionary")
		
		padding = 0
		for each this_cat in docs_in_cat
			scr = get_score(words, this_cat)
			padding = padding - scr
		next
		padding = padding / num_cats
		
		sumprob = 0
		for each this_cat in docs_in_cat
			scr = get_score(words, this_cat) + padding
			this_prob = Exp(scr)
			get_posterior.Add this_cat, this_prob
			sumprob = sumprob + this_prob
		next
		
		if sumprob < 0.0000001 then
			Set get_posterior = Nothing
		else
			for each this_cat in get_posterior
				get_posterior.Item(this_cat) = get_posterior.Item(this_cat) / sumprob
			next
		end if
	end function

	' ########################################################
	' #                                                      #
	' #             get_score                                #
	' #                                                      #
	' ########################################################
	function get_score(words, cat)
		Dim this_word
		get_score = Log(docs_in_cat.Item(cat))
		for each this_word in words
			Dim this_category
			Dim val
			Set this_category = freq_by_word.Item(cat)
			if this_category.Exists(this_word) then
				val = this_category.Item(this_word)
			else
				val = 0
			end if
			
			get_score = get_score + Log((val+1.0)/(words_in_cat.Item(cat) + num_vocs + 0.0))
		next
	end function	
end Class
	
' ####################################################################################
' #                                                                                  #
' #                          Main                                                    #
' #                                                                                  #
' ####################################################################################
Function main()
	Dim arg
	Dim fullpath_train
	
	Set arg = WScript.Arguments
	if arg.Count < 1 then
		WScript.Echo "type (fname_test) | cscript naivebayes.vbs (fname_train) > output.txt"
		Exit Function
	else
		fullpath_train = arg(0)
	end if

	Dim hd
	Set hd = New naivebayes

	hd.train(fullpath_train)
	'hd.print_trained_data()
	hd.test()
	
end Function

main()
