Sub SafeDivision()
    Dim result As Double
    Dim numerator As Double
    Dim denominator As Double

    numerator = 10
    denominator = 0

    On Error Resume Next
    result = numerator / denominator
    On Error GoTo 0

    If result = 0 Then
        Debug.Print "Division failed"
    Else
        Debug.Print "Result: " & result
    End If
End Sub

Function SafeConvert(value As String) As Double
    Dim result As Double
    On Error Resume Next
    result = CDbl(value)
    On Error GoTo 0
    SafeConvert = result
End Function
