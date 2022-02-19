Option Explicit

Sub Create()
    'Audit Trial by Posting

    Application.ScreenUpdating = False

    Rows("2:2").Select
    Selection.Delete Shift:=xlUp

    Dim slips As Collection
    Set slips = New Collection

    Dim i As String: i = 2
    Dim yayoi As yayoi
    Dim debit As Long
    Dim credit As Long
    While Cells(i, 1).Value <> "Total"
        If Cells(i, 1) = "" Then
            GoTo Continue
        End If
        Set yayoi = New yayoi
        yayoi.setDataForSun (i)
        Call slips.Add(yayoi)

        ' debug
         If Cells(i, 10).Value <> "" Then
            debit = debit + Cells(i, 10).Value
            
        Else
            credit = credit + (Cells(i, 11).Value - Cells(i, 11).Value - Cells(i, 11).Value)
        End If

        Continue:
            i = i + 1
    Wend

    Debug.Print debit
    Debug.Print credit

    ' TODO:テンプレートファイルのパスを設定
    Dim src_template As String: src_template = "\yayoi_template.xlsx"
    If Dir(src_template) <> "" Then
        Workbooks.Open src_template
        Workbooks("yayoi_template.xlsx").Activate
    Else
        MsgBox "雛形ファイルが見つかりません。このマクロと同じフォルダに保存して下さい。", vbExclamation
    End If

   i = 2
    For Each yayoi In slips
        Cells(i, 1).Value = "=IF(B2=B1,IF(B2=B3,""2100"",""2101""),""2110"")"
        Cells(i, 2).Value = yayoi.slipNum
'        Cells(i, 3).Value = yayoi(i).financStat
        Cells(i, 4).Value = yayoi.slipDay
        Cells(i, 5).Value = yayoi.debitName
        Cells(i, 6).Value = yayoi.debitSub
        Cells(i, 7).Value = yayoi.debitDep
        Cells(i, 8).Value = yayoi.debitTaxType
        Cells(i, 9).Value = yayoi.debitAmo
        Cells(i, 10).Value = yayoi.debitTax
        Cells(i, 11).Value = yayoi.creditName
        Cells(i, 12).Value = yayoi.creditSub
        Cells(i, 13).Value = yayoi.creditDep
        Cells(i, 14).Value = yayoi.creditTaxType
        Cells(i, 15).Value = yayoi.creditAmo
        Cells(i, 16).Value = yayoi.creditTax
        Cells(i, 17).Value = yayoi.summary
'        Cells(i, 18).Value = yayoi(i).num
'        Cells(i, 19).Value = yayoi(i).settlement
        Cells(i, 20).Value = yayoi.slipType
        Cells(i, 21).Value = yayoi.origin
        Cells(i, 22).Value = yayoi.memo
        Cells(i, 23).Value = yayoi.tag1
        Cells(i, 24).Value = yayoi.tag2
        Cells(i, 25).Value = yayoi.adjustment

        i = i + 1
    Next
 
Application.ScreenUpdating = True

End Sub

