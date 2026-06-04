Sub FormatCell()
    Dim ws As Object

    With ws
        .Name = "DataSheet"
        .Cells(1, 1).Value = "Header"
        .Cells(1, 2).Value = "Amount"
    End With
End Sub

Sub ProcessRecord()
    Dim record As Object

    Set record = New Collection

    With record
        .Add "John", "FirstName"
        .Add "Doe", "LastName"
        .Add 30, "Age"
    End With

    Debug.Print "Done"
End Sub
