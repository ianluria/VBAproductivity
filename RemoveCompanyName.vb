Private Function RemoveCompanyName(inString As String) As String

	Dim stringCompareReturn As Byte
	Dim checkedWord As String
    Dim wordCounter As Integer

	Dim badWords()
	badWords = Split("co,co.,ltd,ltd.,gmbh,spa,s.p.a,inc,inc.", ",")

	Dim wordArrayOfInString() 
	wordArrayOfInString = Split(inString)

	wordCounter = 0

	For Each word In wordArrayOfInString
	
		checkedWord = word
	
		For Each badWord in badWords
			stringCompareReturn = StrComp(word, badWord)
		
			If stringCompareReturn = 0 Then
				checkedWord = ""
				Exit For
			End If 
		
		Next badWord
		 
		wordArrayOfInString(wordCounter) = checkedWord

		wordCounter = wordCounter + 1

	Next word

	RemoveCompanyName = Join(wordArrayOfInString)
End Function



	