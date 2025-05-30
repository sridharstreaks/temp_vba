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

Sub DictionaryIndexExample()

    Dim dict As Object
    Set dict = CreateObject("Scripting.Dictionary")

    ' Add some items to the dictionary
    dict.Add "Key1", "Value1"
    dict.Add "Key2", "Value2"
    dict.Add "Key3", "Value3"

    ' Accessing items using their index (0-based)
    Debug.Print "Item at index 0: " & dict.Item(dict.Keys(0)) 'prints Value1
    Debug.Print "Item at index 1: " & dict.Item(dict.Keys(1)) 'prints Value2
    Debug.Print "Item at index 2: " & dict.Item(dict.Keys(2)) 'prints Value3

    ' Accessing keys and values in a loop
    For i = 0 To dict.Count - 1
        Debug.Print "Key: " & dict.Keys(i) & ", Value: " & dict.Items(i)
    Next i

End Sub
