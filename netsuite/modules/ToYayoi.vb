Option Explicit

'仕訳データをYayoiに格納するプロシージャ
Sub To_yayoi()
    
    '行数取得
    Dim maxRow As Long
    Range("B2").Select
    Selection.End(xlDown).Select
    
    maxRow = Selection.Row
    
    '仕訳データを1行ずつ格納
    Dim Yayoi() As New Yayoi
    ReDim Yayoi(maxRow - 2) As New Yayoi
    
    
    Dim i As Long
    For i = 2 To maxRow
        Yayoi(i - 2).setDate (i)
        
    Next i
   
   Call CreData(Yayoi(), maxRow)
    
End Sub

'インポートファイル作成プロシージャ
Sub CreData(Yayoi() As Yayoi, maxRow As Long)

    Workbooks.Add
    
'    Dim taxDiv As String
'    taxDiv = "課税売上内税10%"
    
    Dim i As Long
    For i = 1 To maxRow - 1
        Cells(i, 1).Value = Yayoi(i - 1).id_Flag
        Cells(i, 2).Value = Yayoi(i - 1).slipNum
'        Cells(i, 3).Value = yayoi(i).financStat
        Cells(i, 4).Value = Yayoi(i - 1).slipDay
        Cells(i, 5).Value = Yayoi(i - 1).debitName
        Cells(i, 6).Value = Yayoi(i - 1).debitSub
        Cells(i, 7).Value = Yayoi(i - 1).debitDep
        Cells(i, 8).Value = Yayoi(i - 1).debitTaxType
        Cells(i, 9).Value = Yayoi(i - 1).debitAmo
        If Yayoi(i - 1).debitTaxType <> "対象外" Then
            Cells(i, 10).Value = Yayoi(i - 1).debitTax
        End If
        Cells(i, 11).Value = Yayoi(i - 1).creditName
        Cells(i, 12).Value = Yayoi(i - 1).creditSub
        Cells(i, 13).Value = Yayoi(i - 1).creditDep
        Cells(i, 14).Value = Yayoi(i - 1).creditTaxType
        Cells(i, 15).Value = Yayoi(i - 1).creditAmo
        If Yayoi(i - 1).creditTaxType <> "対象外" Then
            Cells(i, 16).Value = Yayoi(i - 1).creditTax
        End If
        Cells(i, 17).Value = Yayoi(i - 1).summary
'        Cells(i, 18).Value = yayoi(i).num
'        Cells(i, 19).Value = yayoi(i).settlement
        Cells(i, 20).Value = Yayoi(i - 1).slipType
        Cells(i, 21).Value = Yayoi(i - 1).origin
        Cells(i, 22).Value = Yayoi(i - 1).memo
        Cells(i, 23).Value = Yayoi(i - 1).tag1
        Cells(i, 24).Value = Yayoi(i - 1).tag2
        Cells(i, 25).Value = Yayoi(i - 1).adjustment
        
'        課税売上判定
'        If (Cells(i, 11).Value = "Sales") Then
'          Cells(i, 14).Value = taxDiv
'        End If
'        If (Cells(i, 5).Value = "Sales") Then
'          Cells(i, 8).Value = taxDiv
'        End If
        

    Next i
    
    Columns("I:I").NumberFormatLocal = "###,##0"
    Columns("J:J").NumberFormatLocal = "###,##0"
    Columns("O:O").NumberFormatLocal = "###,##0"
    Columns("P:P").NumberFormatLocal = "###,##0"

End Sub

