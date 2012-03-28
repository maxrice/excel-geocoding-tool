Attribute VB_Name = "mRegex"
' boolean functin tests if regular expression test against string souce
'
' Example: RegExTest("this is a string","[A-Z]") returns False
'          This searches for capital letters in a string
Public Function RegExTest(ByRef source As String, _
                          ByRef test As String) As Boolean
    Dim regex As Object
    Set regex = CreateObject("vbscript.regexp")
        
    regex.Pattern = test
    RegExTest = regex.test(source)
End Function


' Counts the number of matches of regular expression test in source
'
' Example: RegExNumMatches("this is a string","\w+") returns 4
'          The regular expression "\w+" counts words (one or more strings of consecutive letters)
Public Function RegExNumMatches(ByRef source As String, _
                                ByRef test As String) As Integer
    Dim regex As Object
    Set regex = CreateObject("vbscript.regexp")

    
    regex.Pattern = test
    regex.Global = True
    
    Dim match As Object
    Set match = regex.Execute(source)
    
    RegExNumMatches = match.Count
End Function



' Returns a collection object containing all matches of regular expression test against source
Function RegExSubmatches(ByRef source As String, _
                         ByRef test As String) As Object
    Dim regex As Object
    Set regex = CreateObject("vbscript.regexp")
    
    With regex
        .Pattern = test
        .Global = True
    End With
    
    Dim match As Object
    Set match = regex.Execute(source)
    
    If match.Count > 0 Then
        Set RegExSubmatches = match(0).SubMatches
    Else
        Set RegExSubmatches = Nothing
    End If
End Function



' returns a regular expression object after comparing
' test to source
Function RegExMatches(ByRef source As String, _
                      ByRef test As String) As Object
    Dim regex As Object
    Set regex = CreateObject("vbscript.regexp")
    
    Dim match As Object
    
    With regex
        .Pattern = test
        .Global = True
    End With
    
    Set match = regex.Execute(source)
    Set RegExMatches = match
End Function



' Returns the first regular expression match object of comparing regular express test to source
Function RegExMatch(ByRef source As String, _
                      ByRef test As String) As String
    Dim regex As Object
    Set regex = CreateObject("vbscript.regexp")
    
    Dim match As Object
    
    With regex
        .Pattern = test
        .Global = True
    End With
    
    Set match = regex.Execute(source)
    If match.Count > 0 Then
        If match(0).SubMatches.Count > 0 Then
            RegExMatch = match(0).SubMatches(0)
        Else
            RegExMatch = ""
        End If
    Else
        RegExMatch = ""
    End If
End Function




' Returns a string containing only characters in source that match elements in test
'
' Example: RegExValidate("chris gemignani","aeiou") returns "ieiai"
Public Function RegExValidate(ByRef source As String, _
                              ByRef test As String) As String

    Dim s As String
    Dim regex As Object
    Set regex = CreateObject("vbscript.regexp")
    
    With regex
        .Pattern = test
        .Global = True
    End With
    
    Dim matches As Object
    Set matches = regex.Execute(source)
    
    s = ""
    For Each m In matches
        s = s & m
    Next m
    RegExValidate = s
End Function

