Function CleanString(s As String) As String
    Dim result As String
    Dim i As Integer
    Dim ch As String

    result = Trim(s)
    result = UCase(result)

    If Len(result) = 0 Then
        CleanString = ""
        Exit Function
    End If

    If Left(result, 1) = " " Then
        result = Mid(result, 2)
    End If

    CleanString = result
End Function

Sub ProcessData()
    Dim i As Long
    Dim rawValue As String
    Dim cleanValue As String
    Dim count As Long

    count = 0

    For i = 1 To 100
        rawValue = Cells(i, 1).Value

        If rawValue <> "" Then
            cleanValue = CleanString(rawValue)
            Cells(i, 2).Value = cleanValue
            Cells(i, 3).Value = Len(cleanValue)
            count = count + 1
        End If
    Next i

    Range("D1").Value = count
    Debug.Print "Processed " & count & " records"
End Sub

Function FindInRange(searchValue As String, lastRow As Long) As Long
    Dim i As Long

    FindInRange = 0

    For i = 1 To lastRow
        If Cells(i, 1).Value = searchValue Then
            FindInRange = i
            Exit Function
        End If
    Next i
End Function
