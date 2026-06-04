Sub StringOperations()
    Dim s As String
    Dim result As String

    s = "Hello, World!"

    result = Left(s, 5)
    Debug.Print result

    result = Right(s, 6)
    Debug.Print result

    result = Mid(s, 8, 5)
    Debug.Print result

    If InStr(s, "World") > 0 Then
        Debug.Print "Found World"
    End If

    result = Replace(s, "World", "VBA")
    Debug.Print result

    Debug.Print Len(s)
    Debug.Print UCase(s)
    Debug.Print Trim("  hello  ")
End Sub
