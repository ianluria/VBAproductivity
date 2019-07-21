Private Function RemoveSpecialCharacters(inString As String) As String

' Press and hold down the 'ALT' key, and press the 'F11' key.
' Insert a Module in your VBAProject, Microsoft Excel Objects
' To use the function in your workbook you will need to go to a cell in a black column and copy the next formula in a cell. The "inString" in the formula should be replaced with the (same row):
' =RemoveSpecialCharacters(inString)

    Dim stringLength as Integer
    Dim searchChar as String * 1
    Dim foundChar as String * 1

    stringLength = Len(inString)

    ' Array of Strings used to store each character from inString
    Dim cleanedCharArray(stringLength) As String

    For i = 1 To stringLength
        
        searchChar = Mid(inString, i, 1)
        
        Select Case searchChar
            Case "@"                
                foundChar = ""
            Case "#"
                foundChar = ""
            Case "%"
                foundChar = ""
            Case "&"
                foundChar = ""
            Case "?"
                foundChar = ""
            Case "*"
                foundChar = ""
            Case "$"
                foundChar = ""    
            Case Else
                foundChar = searchChar
        End Select
        
        cleanedCharArray(i-1) = foundChar
        
    Next i

    RemoveSpecialCharacters = Join(cleanedCharArray)
    
End Function