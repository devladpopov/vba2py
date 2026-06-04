Function Fibonacci(n As Integer) As Long
    Dim a As Long, b As Long, temp As Long
    Dim i As Integer

    If n <= 0 Then
        Fibonacci = 0
        Exit Function
    ElseIf n = 1 Then
        Fibonacci = 1
        Exit Function
    End If

    a = 0
    b = 1
    For i = 2 To n
        temp = a + b
        a = b
        b = temp
    Next i

    Fibonacci = b
End Function

Function GetGrade(score As Double) As String
    Select Case score
    Case 90, 91, 92, 93, 94, 95, 96, 97, 98, 99, 100
        GetGrade = "A"
    Case 80
        GetGrade = "B"
    Case 70
        GetGrade = "C"
    Case Else
        GetGrade = "F"
    End Select
End Function
