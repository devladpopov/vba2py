Sub TestLoops()
    Dim i As Integer
    Dim total As Long

    ' For loop
    total = 0
    For i = 1 To 10
        total = total + i
    Next i

    ' For with Step
    For i = 10 To 0 Step -2
        Debug.Print i
    Next

    ' Do While (pre-check)
    i = 0
    Do While i < 5
        i = i + 1
    Loop

    ' Do Until (post-check)
    i = 0
    Do
        i = i + 1
    Loop Until i >= 10

    ' While...Wend
    i = 100
    While i > 0
        i = i - 1
    Wend

    Debug.Print "Total: " & total
End Sub
