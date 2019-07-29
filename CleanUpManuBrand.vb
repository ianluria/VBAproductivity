Private Function CleanUpManuBrand(inString As String) As String

' Press and hold down the 'ALT' key, and press the 'F11' key.
' Insert a Module in your VBAProject, Microsoft Excel Objects
' To use the function in your workbook you will need to go to a cell in a black column and copy the next formula in a cell. The "inString" in the formula should be replaced with the (same row):
' =CleanUpManuBrand(inString)

    Dim testChar as String * 1
    Dim cleanedStringOfIllegalChars as String

    ' Remove any illegal characters

    For i = 1 To Len(inString)
        
        testChar = Mid(inString, i, 1)
        
        Select Case testChar
            Case "@"                
               
            Case "#"
                
            Case "%"
                
            Case "?"
                
            Case "*"
                
            Case "$"

            Case "®"

            Case "™"

            Case "&"
                cleanedStringOfIllegalChars = cleanedStringOfIllegalChars + " "
            Case Else
                cleanedStringOfIllegalChars = cleanedStringOfIllegalChars + testChar
        End Select
    Next i

    ' Remove any foreign accents

    Dim fixedChar as String * 1

    For i = 1 To Len(cleanedStringOfIllegalChars)
        
        testChar = Mid(cleanedStringOfIllegalChars, i, 1)
        
        Select Case testChar    
            Case "Š"                ' 138
                fixedChar = "S"
            Case "Ž"                ' 142
                fixedChar = "Z"
            Case "š"                ' 154
                fixedChar = "s"
            Case "ž"                ' 158
                fixedChar = "z"
            Case "Ÿ"                ' 159
                fixedChar = "Y"
            Case "À"                ' 192
                fixedChar = "A"
            Case "Á"                ' 193
                fixedChar = "A"
            Case "Â"                ' 194
                fixedChar = "A"
            Case "Ã"                ' 195
                fixedChar = "A"
            Case "Ä"                ' 196
                fixedChar = "A"
            Case "Å"                ' 197
                fixedChar = "A"
            Case "Ç"                ' 199
                fixedChar = "C"
            Case "È"                ' 200
                fixedChar = "E"
            Case "É"                ' 201
                fixedChar = "E"
            Case "Ê"                ' 202
                fixedChar = "E"
            Case "Ë"                ' 203
                fixedChar = "E"
            Case "Ì"                ' 204
                fixedChar = "I"
            Case "Í"                ' 205
                fixedChar = "I"
            Case "Î"                ' 206
                fixedChar = "I"
            Case "Ï"                ' 207
                fixedChar = "I"
            Case "Ñ"                ' 209
                fixedChar = "N"
            Case "Ò"                ' 210
                fixedChar = "O"
            Case "Ó"                ' 211
                fixedChar = "O"
            Case "Ô"                ' 212
                fixedChar = "O"
            Case "Õ"                ' 213
                fixedChar = "O"
            Case "Ö"                ' 214
                fixedChar = "O"
            Case "Ù"                ' 217
                fixedChar = "U"
            Case "Ú"                ' 218
                fixedChar = "U"
            Case "Û"                ' 219
                fixedChar = "U"
            Case "Ü"                ' 220
                fixedChar = "U"
            Case "Ý"                ' 221
                fixedChar = "Y"
            Case "à"                ' 224
                fixedChar = "a"
            Case "á"                ' 225
                fixedChar = "a"
            Case "â"                ' 226
                fixedChar = "a"
            Case "ã"                ' 227
                fixedChar = "a"
            Case "ä"                ' 228
                fixedChar = "a"
            Case "å"                ' 229
                fixedChar = "a"
            Case "ç"                ' 231
                fixedChar = "c"
            Case "è"                ' 232
                fixedChar = "e"
            Case "é"                ' 233
                fixedChar = "e"
            Case "ê"                ' 234
                fixedChar = "e"
            Case "ë"                ' 235
                fixedChar = "e"
            Case "ì"                ' 236
                fixedChar = "i"
            Case "í"                ' 237
                fixedChar = "i"
            Case "î"                ' 238
                fixedChar = "i"
            Case "ï"                ' 239
                fixedChar = "i"
            Case "ð"                ' 240
                fixedChar = "o"
            Case "ñ"                ' 241
                fixedChar = "n"
            Case "ò"                ' 242
                fixedChar = "o"
            Case "ó"                ' 243
                fixedChar = "o"
            Case "ô"                ' 244
                fixedChar = "o"
            Case "õ"                ' 245
                fixedChar = "o"
            Case "ö"                ' 246
                fixedChar = "o"
            Case "ù"                ' 249
                fixedChar = "u"
            Case "ú"                ' 250
                fixedChar = "u"
            Case "û"                ' 251
                fixedChar = "u"
            Case "ü"                ' 252
                fixedChar = "u"
            Case "ý"                ' 253
                fixedChar = "y"
            Case "ÿ"                ' 255
                fixedChar = "y"
            Case Else
                fixedChar = testChar
        End Select
         
        CleanUpManuBrand = CleanUpManuBrand + fixedChar

    Next i

    ' Remove any company names

    Dim badWords as Variant
	badWords = Split("co,co.,ltd,ltd.,gmbh,spa,s.p.a.,inc,inc.,sa,s.a,s.a.,sl,s.l, s.l.,pvt", ",")

	Dim wordArrayOfInString as Variant 
	wordArrayOfInString = Split(CleanUpManuBrand)

    CleanUpManuBrand = ""

	Dim stringCompareReturn As Integer
    Dim isBadWord As Boolean

	For Each word In wordArrayOfInString

        isBadWord = False
	
		For Each badWord in badWords
			stringCompareReturn = StrComp(word, badWord, 1)
		
			If stringCompareReturn = 0 Then
				isBadWord = True
				Exit For
			End If 
		
		Next badWord

        If isBadWord = False Then 
		    CleanUpManuBrand = CleanUpManuBrand + " " + word
        End If

	Next word

	CleanUpManuBrand = Trim(UCase(CleanUpManuBrand))
    
End Function