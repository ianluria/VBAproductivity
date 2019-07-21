Private Function CheckStringChar(inString As String) As String

    ' CheckStringChar(inString)
    ' Returns its passed agrument, but with exchanged characters
    ' Function created 7/08/2003 by Stanley D. Grom, Jr.
    ' Greatly Modified by Ian Luria 6/3/19

    ' Press and hold down the 'ALT' key, and press the 'F11' key.
    ' Insert a Module in your VBAProject, Microsoft Excel Objects
    ' To use the function in your workbook you will need to go to a cell in a black column and copy the next formula in a cell. The "inString" in the formula should be replaced with the (same row):
    ' =CheckStringChar(inString)

    Dim stringLength as Integer
    Dim searchChar as String * 1 
    Dim foundChar as String * 1

    stringLength = Len(inString)

    Dim cleanedCharArray(stringLength) As String 

    For i = 1 To stringLength
        
        searchChar = Mid(inString, i, 1)
        
        Select Case searchChar
            Case "Š"                ' 138
                foundChar = "S"
            Case "Ž"                ' 142
                foundChar = "Z"
            Case "š"                ' 154
                foundChar = "s"
            Case "ž"                ' 158
                foundChar = "z"
            Case "Ÿ"                ' 159
                foundChar = "Y"
            Case "À"                ' 192
                foundChar = "A"
            Case "Á"                ' 193
                foundChar = "A"
            Case "Â"                ' 194
                foundChar = "A"
            Case "Ã"                ' 195
                foundChar = "A"
            Case "Ä"                ' 196
                foundChar = "A"
            Case "Å"                ' 197
                foundChar = "A"
            Case "Ç"                ' 199
                foundChar = "C"
            Case "È"                ' 200
                foundChar = "E"
            Case "É"                ' 201
                foundChar = "E"
            Case "Ê"                ' 202
                foundChar = "E"
            Case "Ë"                ' 203
                foundChar = "E"
            Case "Ì"                ' 204
                foundChar = "I"
            Case "Í"                ' 205
                foundChar = "I"
            Case "Î"                ' 206
                foundChar = "I"
            Case "Ï"                ' 207
                foundChar = "I"
            Case "Ñ"                ' 209
                foundChar = "N"
            Case "Ò"                ' 210
                foundChar = "O"
            Case "Ó"                ' 211
                foundChar = "O"
            Case "Ô"                ' 212
                foundChar = "O"
            Case "Õ"                ' 213
                foundChar = "O"
            Case "Ö"                ' 214
                foundChar = "O"
            Case "Ù"                ' 217
                foundChar = "U"
            Case "Ú"                ' 218
                foundChar = "U"
            Case "Û"                ' 219
                foundChar = "U"
            Case "Ü"                ' 220
                foundChar = "U"
            Case "Ý"                ' 221
                foundChar = "Y"
            Case "à"                ' 224
                foundChar = "a"
            Case "á"                ' 225
                foundChar = "a"
            Case "â"                ' 226
                foundChar = "a"
            Case "ã"                ' 227
                foundChar = "a"
            Case "ä"                ' 228
                foundChar = "a"
            Case "å"                ' 229
                foundChar = "a"
            Case "ç"                ' 231
                foundChar = "c"
            Case "è"                ' 232
                foundChar = "e"
            Case "é"                ' 233
                foundChar = "e"
            Case "ê"                ' 234
                foundChar = "e"
            Case "ë"                ' 235
                foundChar = "e"
            Case "ì"                ' 236
                foundChar = "i"
            Case "í"                ' 237
                foundChar = "i"
            Case "î"                ' 238
                foundChar = "i"
            Case "ï"                ' 239
                foundChar = "i"
            Case "ð"                ' 240
                foundChar = "o"
            Case "ñ"                ' 241
                foundChar = "n"
            Case "ò"                ' 242
                foundChar = "o"
            Case "ó"                ' 243
                foundChar = "o"
            Case "ô"                ' 244
                foundChar = "o"
            Case "õ"                ' 245
                foundChar = "o"
            Case "ö"                ' 246
                foundChar = "o"
            Case "ù"                ' 249
                foundChar = "u"
            Case "ú"                ' 250
                foundChar = "u"
            Case "û"                ' 251
                foundChar = "u"
            Case "ü"                ' 252
                foundChar = "u"
            Case "ý"                ' 253
                foundChar = "y"
            Case "ÿ"                ' 255
                foundChar = "y"
            Case Else
                foundChar = searchChar
        End Select
        
        cleanedCharArray(i - 1) = foundChar

    Next i  
    
    CheckStringChar = Join(cleanedCharArray)
    
End Function