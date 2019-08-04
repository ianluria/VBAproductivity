Function simpleCellRegex(userInput As String) As String

    Dim regEx As New RegExp
    Dim badCharacterPattern As String

    Dim aPattern As String
    Dim ePattern As String
    Dim oPattern As String
    Dim iPattern As String
    Dim uPattern As String
    Dim yPattern As String
    Dim abbreviationPattern As String

    badCharacterPattern = "[@!#$%^&*]"
    abbreviationPattern = "[,.]+com|[\s,.]g[.]?m[.]?b[.]?h[.]?|[\s,.]c[.]?o[.]?|[\s,.]l[.]?t[.]?d[.]?|[\s,.]s[.]?p[.]?a[.]?|[\s,.]s[.]?a[.]?|[\s,.]i[.]?n[.]?c[.]?|[\s,.]s[.]?l[.]?|[\s,.]p[.]?v[.]?t[.]?"
    aPattern = "[ÀÁÂÃÄÅáâãäå]"
    ePattern = "[ÈÉÊËëêéè]"
    oPattern = "[ÒÓÔÕÖðòóôõö]"
    iPattern = "[ÌÍÎÏìíîï]"
    uPattern = "[ÙÚÛÜùúûü]"
    otherPattern"[ŸÝýÿÇçŽžÑñŠš]"
    yPattern = "[ŸÝýÿ]"
    cPattern = "[Çç]"
    zPattern = "[Žž]"
    nPattern = "[Ññ]"
    sPattern = "[Šš]"

    With regEx
        .Global = True
        .MultiLine = True
        .IgnoreCase = False
        .Pattern = badCharacterPattern
    End With

    If regEx.test(userInput) Then
        simpleCellRegex = regEx.Replace(userInput, "")













    Else 
        simpleCellRegex = "No"
    End If

End Function    
        
