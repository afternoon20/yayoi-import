Option Explicit

Sub Create()
    Application.ScreenUpdating = False
    
    Cells(8, 1).Select
    Selection.End(xlDown).Select
    Dim journal_count As Long: journal_count = ActiveCell.Row
    Dim Yayoi() As New Yayoi
    ReDim Yayoi(journal_count - 2) As New Yayoi
    
    Dim i As Long: i = 8
    Dim j As Long: j = 0
    While Cells(i, 1).Value <> ""
        Yayoi(j).slipNum = Cells(i, 1).Value
        Yayoi(j).slipDay = Cells(i, 5).Value
'        借方科目の場合
        If Cells(i, 9).Value <> "" Then
            Yayoi(j).debitName = Cells(i, 2).Value
            Yayoi(j).debitAmo = Cells(i, 9).Value
        End If
        
'        貸方科目の場合
        If Cells(i, 10).Value <> "" Then
            Yayoi(j).creditName = Cells(i, 2).Value
            Yayoi(j).creditAmo = Cells(i, 10).Value
        End If
        Yayoi(j).summary = Cells(i, 8).Value
        
        j = j + 1
        i = i + 1
    
    Wend
    
    Dim src_template As String: src_template = "C:\Users\shimada\Desktop\yayoi_template.xlsx"
    If Dir(src_template) <> "" Then
        Workbooks.Open src_template
        Workbooks("yayoi_template.xlsx").Activate
    Else
        MsgBox "雛形ファイルが見つかりません。このマクロと同じフォルダに保存して下さい。", vbExclamation
    End If
    
    For i = 2 To UBound(Yayoi) - 2
'        Cells(i, 1).Value = Yayoi(i - 2).id_Flag
        Cells(i, 1).Value = "=IF(B1871=B1870,IF(B1871=B1872,""2100"",""2101""),""2110"")"
        Cells(i, 2).Value = Yayoi(i - 2).slipNum
'        Cells(i, 3).Value = yayoi(i).financStat
        Cells(i, 4).Value = Yayoi(i - 2).slipDay
        Cells(i, 5).Value = Yayoi(i - 2).debitName
        Cells(i, 6).Value = Yayoi(i - 2).debitSub
        Cells(i, 7).Value = Yayoi(i - 2).debitDep
        Cells(i, 8).Value = Yayoi(i - 2).debitTaxType
        Cells(i, 9).Value = Yayoi(i - 2).debitAmo
        Cells(i, 10).Value = Yayoi(i - 2).debitTax
        Cells(i, 11).Value = Yayoi(i - 2).creditName
        Cells(i, 12).Value = Yayoi(i - 2).creditSub
        Cells(i, 13).Value = Yayoi(i - 2).creditDep
        Cells(i, 14).Value = Yayoi(i - 2).creditTaxType
        Cells(i, 15).Value = Yayoi(i - 2).creditAmo
        Cells(i, 16).Value = Yayoi(i - 2).creditTax
        Cells(i, 17).Value = Yayoi(i - 2).summary
'        Cells(i, 18).Value = yayoi(i).num
'        Cells(i, 19).Value = yayoi(i).settlement
        Cells(i, 20).Value = Yayoi(i - 2).slipType
        Cells(i, 21).Value = Yayoi(i - 2).origin
        Cells(i, 22).Value = Yayoi(i - 2).memo
        Cells(i, 23).Value = Yayoi(i - 2).tag1
        Cells(i, 24).Value = Yayoi(i - 2).tag2
        Cells(i, 25).Value = Yayoi(i - 2).adjustment
    Next i
    
    
    
    Dim slipNum As Long
    slipNum = Cells(2, 2).Value
    Cells(2, 1).Value = "2110"

    j = 1
    For i = 3 To UBound(Yayoi) - 1
        If Cells(i, 2).Value = slipNum Then
            Cells(i, 1).Value = 2100
            j = j + 1

        Else
            Cells(i, 1).Value = 2110
            slipNum = slipNum + 1

            If j = 1 Then
                Cells(i - 1, 1).Value = 2111
            Else
              Cells(i - 1, 1).Value = 2101
            End If
            j = 1
        End If

    Next i
    
    Application.ScreenUpdating = True
    

End Sub

