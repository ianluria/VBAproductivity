Private Function RemoveCompanyName(InString As String)

	Dim ArrayOfWords() As String 'varaint?
	Dim BadWords() As String
	Dim StringCompareReturn As Integer
	Dim CheckedWord As String

	BadWords = Split("co,co.,ltd,ltd.,gmbh,spa,s.p.a,inc,inc.", ",")

	ArrayOfWords = Split(InString)

	For Each Word In ArrayOfWords
	
		CheckedWord = Word
	
		For Each BadWord in BadWords
			StringCompareReturn = StrComp(Word, BadWord)
		
			If StringCompareReturn = 0 Then
				CheckedWord = ""
				Exit For
		
		Next BadWord
		' May need to fix this
		RemoveCompanyName = RemoveCompanyName & CheckedWord
	Next Word

End Function



	