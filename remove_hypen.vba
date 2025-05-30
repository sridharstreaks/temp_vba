Function RemoveCharsAfterHyphen(strText As String) As String
    Dim hyphenPos As Integer
    
    ' Find the position of the hyphen
    hyphenPos = InStr(1, strText, "-", vbTextCompare)
    
    ' If no hyphen is found, return the original string
    If hyphenPos = 0 Then
        RemoveCharsAfterHyphen = strText
        Exit Function
    End If
    
    ' Return the substring before the hyphen
    RemoveCharsAfterHyphen = Left(strText, hyphenPos - 1)
End Function

Sub ExampleUsage()
    Dim myString As String
    Dim result As String
    
    myString = "123-456-789"
    result = RemoveCharsAfterHyphen(myString)
    Debug.Print result ' Output: 123
    
    myString = "ABC-DEF"
    result = RemoveCharsAfterHyphen(myString)
    Debug.Print result ' Output: ABC
    
    myString = "Just some text"
    result = RemoveCharsAfterHyphen(myString)
    Debug.Print result ' Output: Just some text (no hyphen found)
End Sub
