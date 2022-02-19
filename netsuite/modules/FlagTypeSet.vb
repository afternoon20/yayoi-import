Option Explicit

Sub flagTypeSet()
    Dim slipNum As Long: slipNum = Cells(2, 2).Value
    Cells(2, 1).Value = "2110"

    Dim i As Long: i = 3
    Dim j As Long: j = 1
    While Cells(i, 2).Value <> ""
        If Cells(i, 2).Value = slipNum Then
            Cells(i, 1).Value = 2100
            j = j + 1

        Else
            Cells(i, 1).Value = 2110
            slipNum = slipNum + 1

            If j = 1 Then
                Cells(i - 1, 1).Value = 2111
                Cells(i, 20).Value = 1
            Else
                Cells(i - 1, 1).Value = 2101
            End If
            j = 1
        End If
        i = i + 1
    Wend
    i = i - 1
    If Cells(i, 2).Value = Cells(i - 1, 2).Value Then
        Cells(i, 1).Value = 2101
    Else
        Cells(i, 1).Value = 2111
    End If
    
    

End Sub
