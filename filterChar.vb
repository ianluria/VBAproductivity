Private Function CheckStringCHAR(InString) As String

' CheckStringCHAR(InString)
' Returns its passed agrument, but with exchanged European? characters
' Function created 7/08/2003 by Stanley D. Grom, Jr.
' Modified by Ian Luria 6/3/19

' Press and hold down the 'ALT' key, and press the 'F11' key.
' Insert a Module in your VBAProject, Microsoft Excel Objects
' To use the function in your workbook you will need to go to a cell in a black column and copy the next formula in a cell. The "InString" in the formula should be replaced with the (same row):
' =CheckStringCHAR(InString)


CheckStringCHAR = ""
StringLength = Len(InString)

For i = 1 To StringLength
    
    SearchCHAR = Mid(InString, i, 1)
    
    Select Case SearchCHAR
        Case "Š"                ' 138
            FoundCHAR = "S"
        Case "Ž"                ' 142
            FoundCHAR = "Z"
        Case "š"                ' 154
            FoundCHAR = "s"
        Case "ž"                ' 158
            FoundCHAR = "z"
        Case "Ÿ"                ' 159
            FoundCHAR = "Y"
        Case "À"                ' 192
            FoundCHAR = "A"
        Case "Á"                ' 193
            FoundCHAR = "A"
        Case "Â"                ' 194
            FoundCHAR = "A"
        Case "Ã"                ' 195
            FoundCHAR = "A"
        Case "Ä"                ' 196
            FoundCHAR = "A"
        Case "Å"                ' 197
            FoundCHAR = "A"
        Case "Ç"                ' 199
            FoundCHAR = "C"
        Case "È"                ' 200
            FoundCHAR = "E"
        Case "É"                ' 201
            FoundCHAR = "E"
        Case "Ê"                ' 202
            FoundCHAR = "E"
        Case "Ë"                ' 203
            FoundCHAR = "E"
        Case "Ì"                ' 204
            FoundCHAR = "I"
        Case "Í"                ' 205
            FoundCHAR = "I"
        Case "Î"                ' 206
            FoundCHAR = "I"
        Case "Ï"                ' 207
            FoundCHAR = "I"
        Case "Ñ"                ' 209
            FoundCHAR = "N"
        Case "Ò"                ' 210
            FoundCHAR = "O"
        Case "Ó"                ' 211
            FoundCHAR = "O"
        Case "Ô"                ' 212
            FoundCHAR = "O"
        Case "Õ"                ' 213
            FoundCHAR = "O"
        Case "Ö"                ' 214
            FoundCHAR = "O"
        Case "Ù"                ' 217
            FoundCHAR = "U"
        Case "Ú"                ' 218
            FoundCHAR = "U"
        Case "Û"                ' 219
            FoundCHAR = "U"
        Case "Ü"                ' 220
            FoundCHAR = "U"
        Case "Ý"                ' 221
            FoundCHAR = "Y"
        Case "à"                ' 224
            FoundCHAR = "a"
        Case "á"                ' 225
            FoundCHAR = "a"
        Case "â"                ' 226
            FoundCHAR = "a"
        Case "ã"                ' 227
            FoundCHAR = "a"
        Case "ä"                ' 228
            FoundCHAR = "a"
        Case "å"                ' 229
            FoundCHAR = "a"
        Case "ç"                ' 231
            FoundCHAR = "c"
        Case "è"                ' 232
            FoundCHAR = "e"
        Case "é"                ' 233
            FoundCHAR = "e"
        Case "ê"                ' 234
            FoundCHAR = "e"
        Case "ë"                ' 235
            FoundCHAR = "e"
        Case "ì"                ' 236
            FoundCHAR = "i"
        Case "í"                ' 237
            FoundCHAR = "i"
        Case "î"                ' 238
            FoundCHAR = "i"
        Case "ï"                ' 239
            FoundCHAR = "i"
        Case "ð"                ' 240
            FoundCHAR = "o"
        Case "ñ"                ' 241
            FoundCHAR = "n"
        Case "ò"                ' 242
            FoundCHAR = "o"
        Case "ó"                ' 243
            FoundCHAR = "o"
        Case "ô"                ' 244
            FoundCHAR = "o"
        Case "õ"                ' 245
            FoundCHAR = "o"
        Case "ö"                ' 246
            FoundCHAR = "o"
        Case "ù"                ' 249
            FoundCHAR = "u"
        Case "ú"                ' 250
            FoundCHAR = "u"
        Case "û"                ' 251
            FoundCHAR = "u"
        Case "ü"                ' 252
            FoundCHAR = "u"
        Case "ý"                ' 253
            FoundCHAR = "y"
        Case "ÿ"                ' 255
            FoundCHAR = "y"
        Case Else
            FoundCHAR = SearchCHAR
    End Select
    
    CheckStringCHAR = CheckStringCHAR & FoundCHAR
    
Next i
    
End Function