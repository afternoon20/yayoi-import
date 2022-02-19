Option Explicit

    '識別フラグ
    Public id_Flag As Long 
    '伝票No(管理用)
    Public slipNum As Long
    '決算
    Public financStat As Long
    '日付
    Public slipDay As Date
    '借方勘定科目
    Public debitName As String
    '借方補助科目
    Public debitSub As String
    '借方部門
    Public debitDep As String
    '借方税区分
    Public debitTaxType As String
    '借方金額
    Public debitAmo As Single
    '借方税金額
    Public debitTax As Single
    '貸方勘定科目
    Public creditName As String
    '貸方補助科目
    Public creditSub As String
    '貸方部門
    Public creditDep As String
    '貸方税区分
    Public creditTaxType As String
    '貸方金額
    Public creditAmo As Single
    '貸方税金額
    Public creditTax As Single
    '摘要
    Public summary As String
    '番号
    Public num As Long
    '期日
    Public settlement As Date
    'タイプ（仕訳データの場合は「0」、振伝は「3」）
    Public slipType As Long
    '生成元
    Public origin As String
    '仕訳メモ
    Public memo As String
    '付箋1
    Public tag1 As String
    '付箋2
    Public tag2 As String
    '調整（noと記入）
    Public adjustment As String
 
    Private Sub Class_Initialize()
        Me.debitTaxType = "対象外"
        Me.creditTaxType = "対象外"
        Me.adjustment = "no"
        Me.slipType = 3
    End Sub

    Private Sub Class_Terminate()
    End Sub
    
    Public Sub setDate(i As Long)
        Me.id_Flag = Cells(i, 1).Value
        Me.slipNum = Cells(i, 2).Value
        Me.financStat = Cells(i, 3).Value
        Me.slipDay = Cells(i, 4).Value
        Me.debitName = Cells(i, 5).Value
        Me.debitSub = Cells(i, 6).Value
        Me.debitDep = Cells(i, 7).Value
        Me.debitTaxType = Cells(i, 8).Value
        Me.debitAmo = Cells(i, 9).Value
        Me.debitTax = Cells(i, 10).Value
        Me.creditName = Cells(i, 11).Value
        Me.creditSub = Cells(i, 12).Value
        Me.creditDep = Cells(i, 13).Value
        Me.creditTaxType = Cells(i, 14).Value
        Me.creditAmo = Cells(i, 15).Value
        Me.creditTax = Cells(i, 16).Value
        Me.summary = Cells(i, 17).Value
        Me.num = Cells(i, 18).Value
        Me.settlement = Cells(i, 19).Value
        Me.slipType = Cells(i, 20).Value
        Me.origin = Cells(i, 21).Value
        Me.memo = Cells(i, 22).Value
        Me.tag1 = Cells(i, 23).Value
        Me.tag2 = Cells(i, 24).Value
    End Sub

    ' TODO:テンプレートファイルに直接設定できるようにする。
    Public Sub setDisplay(i As Long)

    end Sub

