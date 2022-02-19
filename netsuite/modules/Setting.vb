Option Explicit

Sub Setting()

'    Rows("1:6").Select
'    Selection.Delete Shift:=xlUp

    Application.ScreenUpdating = False

    Cells.Select
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    
'    勘定科目設定
    Dim account_name As String: account_name = Cells(8, 1).Value
    Dim i As Long: i = 9
    
    While Cells(i, 10).Value <> ""
        If Cells(i, 1).Value = "" Then
            Cells(i, 1).Value = account_name
        Else
            account_name = Cells(i, 1).Value
        End If
        
'        金額の切り上げ
        If Cells(i, 8).Value <> "" Then
            Cells(i, 8).Value = WorksheetFunction.RoundUp(Cells(i, 8).Value, 0)
        End If
        
        If Cells(i, 9).Value <> "" Then
            Cells(i, 9).Value = WorksheetFunction.RoundUp(Cells(i, 9).Value, 0)
        End If
        
'        日付の整形
        Cells(i, 4).Value = Format(Cells(i, 4).Value, "yyyy/mm/dd")
        
        i = i + 1
    Wend
    
'    不要な行削除：合計、空白行
    i = 8
    While Cells(i, 1).Value <> ""
        If (Cells(i, 8).Value <> "" And Cells(i, 9).Value <> "") Or (Cells(i, 8).Value = "" And Cells(i, 9).Value = "") Then
            Rows(i).Delete
            i = i - 1
       End If
        
        i = i + 1
    Wend
    
    Columns("A:A").Select
    With Selection
        .HorizontalAlignment = xlGeneral
        .VerticalAlignment = xlBottom
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = -1
        .ShrinkToFit = False
        .ReadingOrder = xlContext
    End With
    
'    ソート：Document Number
    Range("A7:J7").Select
    Selection.AutoFilter
 
    If ActiveSheet.AutoFilterMode = False Then
        Selection.AutoFilter
    End If
    
    ActiveWorkbook.Worksheets(1).AutoFilter.Sort. _
        SortFields.Clear
    ActiveWorkbook.Worksheets(1).AutoFilter.Sort. _
        SortFields.Add2 Key:=Range("E7"), SortOn:=xlSortOnValues, Order:= _
        xlAscending, DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets(1).AutoFilter.Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With

'    仕訳番号付与
    Range("A1:J6").Select
    With Selection
        .HorizontalAlignment = xlGeneral
        .VerticalAlignment = xlBottom
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    With Selection
        .HorizontalAlignment = xlGeneral
        .VerticalAlignment = xlBottom
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    With Selection
        .HorizontalAlignment = xlGeneral
        .VerticalAlignment = xlBottom
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    With Selection
        .HorizontalAlignment = xlGeneral
        .VerticalAlignment = xlBottom
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    With Selection
        .HorizontalAlignment = xlGeneral
        .VerticalAlignment = xlBottom
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    With Selection
        .HorizontalAlignment = xlGeneral
        .VerticalAlignment = xlBottom
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    Selection.UnMerge
    Columns("A:A").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    
    
    Range("B7").Select
    Selection.Copy
    Range("A7").Select
    ActiveSheet.Paste
    Application.CutCopyMode = False
    ActiveCell.FormulaR1C1 = "Number"
    
    i = 8
    Dim j As Long: j = 1
    Dim document_number As String: document_number = Cells(i, 6).Value
    Cells(i, 1).Value = j
    Dim before_date As Date
    
    i = 9
    
    While Cells(i, 5).Value <> ""
        If Cells(i, 6).Value = document_number Then
            Cells(i, 1).Value = j
        Else
            j = j + 1
            document_number = Cells(i, 6).Value
            Cells(i, 1).Value = j
        End If
        
        
        i = i + 1
    Wend
    
    Columns("A:A").Select
    Range("A7").Activate
    With Selection
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlBottom
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
    End With
    
'    再ソート
    Range("A7:J7").Select
    Selection.AutoFilter
 
    If ActiveSheet.AutoFilterMode = False Then
        Selection.AutoFilter
    End If
    
    ActiveWorkbook.Worksheets(1).AutoFilter.Sort. _
        SortFields.Clear
    ActiveWorkbook.Worksheets(1).AutoFilter.Sort. _
        SortFields.Add2 Key:=Range("F7"), SortOn:=xlSortOnValues, Order:= _
        xlAscending, DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets(1).AutoFilter.Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    
    
'    勘定科目の数字削除
'    Dim reg As Object
'    Set reg = CreateObject("VBScript.RegExp")
'
'    With reg
'        .Pattern = "\d{5}..."
'        .IgnoreCase = False
'        .Global = True
'    End With
'
'    i = 8
'
'    While Cells(i, 2).Value <> ""
'        Dim after_account_name As String
'        after_account_name = reg.Replace(Cells(i, 2).Value, "")
'         Cells(i, 2).Value = after_account_name
'        i = i + 1
'    Wend

    
    
   Application.ScreenUpdating = True

End Sub
