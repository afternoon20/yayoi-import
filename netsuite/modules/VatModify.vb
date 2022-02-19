Option Explicit

Sub vatModify()

'TODO:消費税の科目設定
Const TAX_NAME As String = " 仮払消費税 Consumption Tax Receivable"
Const TAX_TYPE As String = "課対仕入内"

Dim i As Long: i = 1
Dim j As Long: j = 1
Dim slipNum As Long
Dim taxValue As Long
While Cells(i, 2).Value <> ""
    If Cells(i, 5).Value = TAX_NAME Then
        slipNum = Cells(i, 2).Value
        taxValue = Cells(i, 9).Value

'        以降の行のいずれかが本体価格か判定
        If Cells(i + 1, 2).Value = slipNum Then
            j = i + 1
            While Cells(j, 2).Value = slipNum
                If Cells(j, 9).Value = taxValue * 10 Or Cells(j, 9).Value = taxValue * 100 / 8 Then
                    Debug.Print (slipNum)
                
                    If Cells(j, 9).Value = taxValue * 10 Then
                        Cells(j, 8).Value = TAX_TYPE & "10%"
                        Cells(j, 9).Value = Cells(j, 9).Value + taxValue
                        Cells(j, 10).Value = taxValue
                    
                    
                    ElseIf Cells(j, 9).Value = taxValue * 100 / 8 Then
                        Cells(j, 8).Value = TAX_TYPE & "8%"
                        Cells(j, 9).Value = Cells(j, 9).Value + taxValue
                        Cells(j, 10).Value = taxValue
                    
                    End If
                End If
                j = j + 1
            Wend
        
        '        以前の行のいずれかが本体価格か判定
        ElseIf Cells(i - 1, 2).Value = slipNum Then
            j = i + 1
            While Cells(j, 2).Value = slipNum
                If Cells(j, 9).Value = taxValue * 10 Or Cells(j, 9).Value = taxValue * 100 / 8 Then
                    Debug.Print (slipNum)
                
                    If Cells(j, 9).Value = taxValue * 10 Then
                        Cells(j, 8).Value = TAX_TYPE & "10%"
                        Cells(j, 9).Value = Cells(j, 9).Value + taxValue
                        Cells(j, 10).Value = taxValue
                    
                    
                    ElseIf Cells(j, 9).Value = taxValue * 100 / 8 Then
                        Cells(j, 8).Value = TAX_TYPE & "8%"
                        Cells(j, 9).Value = Cells(j, 9).Value + taxValue
                        Cells(j, 10).Value = taxValue
                    
                    End If
                End If
                j = j - 1
            Wend
        End If
        
        Rows(i).Delete
        i = i - 1
    End If
    i = i + 1
        
Wend

Call flagTypeSet.flagTypeSet

End Sub
