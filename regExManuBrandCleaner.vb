Function simpleCellRegex(Myrange As Range) As String

    Dim regEx As New RegExp
    Dim badLetterPattern As String
    Dim userInput As String
    Dim stringReplace As String
    Dim aPattern As String
    Dim ePattern As String
    Dim oPattern As String
    Dim iPattern As String
    Dim UPattern As String
    Dim yPattern As String
    Dim miscPattern As String

    userInput = Myrange.Value

    badLetterPattern = "[aaaaeeeeeiiiiooouuuu@!#$%^&*]"
    aPattern
    ePattern
    oPattern
    iPattern
    UPattern
    yPattern
    miscPattern

    With regEx
        .Global = True
        .MultiLine = True
        .IgnoreCase = False
        .Pattern = badLetterPattern
    End With

    If regEx.test(userInput) Then
        regEx.Pattern = aPattern
        
