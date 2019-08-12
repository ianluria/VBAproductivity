Function simpleCellRegex(userInput As String) As String

    Dim regEx As New RegExp

    Dim aPattern As String
        aPattern = "[ÀÁÂÃÄÅáâãäå]"
    Dim ePattern As String
        ePattern = "[ÈÉÊËëêéè]"
    Dim iPattern As String
        iPattern = "[ÌÍÎÏìíîï]"
    Dim oPattern As String
        oPattern = "[ÒÓÔÕÖðòóôõö]"
    Dim uPattern As String
        uPattern = "[ÙÚÛÜùúûü]"
    Dim yPattern As String
        yPattern = "[ŸÝýÿ]"
    Dim cPattern As String
        cPattern = "[Çç]"
    Dim zPattern As String
        zPattern = "[Žž]"
    Dim nPattern As String
        nPattern = "[Ññ]"
    Dim sPattern As String
        sPattern = "[Šš]"

    Dim vowelPattern As String
        vowelPattern = "[ÀÁÂÃÄÅáâãäåÈÉÊËëêéèÒÓÔÕÖðòóôõöÌÍÎÏìíîïÙÚÛÜùúûüŸÝýÿÇçŽžÑñŠš]"
    Dim otherPattern As String
        otherPattern = "[ŸÝýÿÇçŽžÑñŠš]"

    Dim badCharacterPattern As String
        badCharacterPattern = "[@!#$%^&*]"

    Dim abbreviationPattern As String
        abbreviationPattern = "[,.]+com|[\s,.]g[.]?m[.]?b[.]?h[.]?|[\s,.]c[.]?o[.]?|[\s,.]l[.]?t[.]?d[.]?|[\s,.]s[.]?p[.]?a[.]?|[\s,.]s[.]?a[.]?|[\s,.]i[.]?n[.]?c[.]?|[\s,.]s[.]?l[.]?|[\s,.]p[.]?v[.]?t[.]?"

    Dim replaceCharacters As String
        replaceCharacters = ",A,E,I,O,U,Y,C,Z,N,S, "
    Dim replaceCharacterArray() As String
        replaceCharacterArray = Split(replaceCharacters, ",")

    Dim arrayOfRegexes As Variant
        arrayOfRegexes = Array(badCharacterPattern,aPattern,ePattern,iPattern,oPattern,uPattern,yPattern,cPattern,zPattern,nPattern,sPattern,abbreviationPattern)

    With regEx
        .Global = True
        .MultiLine = True
        .IgnoreCase = False
        .Pattern = vowelPattern
    End With

    Dim accentedVowelPresent As Boolean
    Dim otherAccentPresent As Boolean

    simpleCellRegex = userInput

    ' Test for presence of accented character
    If regEx.test(simpleCellRegex) Then
        accentedVowelPresent = True
        regEx.Pattern = otherPattern
        ' Test for less common accents 
        If regEx.test(simpleCellRegex) Then      
            otherAccentPresent = True
        End If
    End If

    Dim replace As Boolean

    For i = LBound(arrayOfRegexes) To UBound(arrayOfRegexes)
        
        regEx.Pattern = arrayOfRegexes(i)
        replace = False

        If i = LBound(arrayOfRegexes) OR i = UBound(arrayOfRegexes) Then
            replace = True
        Else
            If i < 6 Then
                If accentedVowelPresent Then
                    replace = True
                End If
            Else
                If otherAccentPresent Then
                    replace = True
                End If
            End If    
        End If

        If replace Then
            simpleCellRegex = regEx.Replace(simpleCellRegex, replaceCharacterArray(i))
        End If    

    Next i
    
    simpleCellRegex = UCase(simpleCellRegex)
    simpleCellRegex = Trim(simpleCellRegex)

End Function    