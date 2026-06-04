Sub SumColumn()
    Dim total As Double
    Dim lastRow As Long
    Dim i As Long

    lastRow = 100
    total = 0

    For i = 1 To lastRow
        If Cells(i, 1).Value <> "" Then
            total = total + Cells(i, 1).Value
        End If
    Next i

    Range("B1").Value = total
    Debug.Print "Total: " & total
End Sub

Sub CopyRange()
    Dim srcRow As Long
    Dim destRow As Long

    destRow = 1
    For srcRow = 1 To 50
        If Cells(srcRow, 1).Value > 0 Then
            Cells(destRow, 3).Value = Cells(srcRow, 1).Value
            Cells(destRow, 4).Value = Cells(srcRow, 2).Value
            destRow = destRow + 1
        End If
    Next srcRow
End Sub
