Sub 事前入力csv_main_印刷時でmsgbox無し()
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic '計算自動
    
    'アクティブ
    Worksheets("ピッキング表").Activate
    Worksheets("ピッキング表").Select
    kakunin_ship = Range("D6")

    Worksheets("ピッキング表").Unprotect   '保護解除
    'A1セルから空白セルまでdowhileで回し、その領域内をcsv出力する
    
    'Application.CalculateFull   '全シート再計算
    
    '非表示
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    
    Call センター数量取得
    
    'アクティブ
    Worksheets("事前出力用").Activate
    Worksheets("事前出力用").Select
    Range("A1").Select
    
    'ワークシート
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets("事前出力用")

    'G2 出荷先
    Dim Customer_name As String
    Customer_name = Range("G2")

    'F1出荷日
    Dim ship_date As Date
    ship_date = Range("D1")

    'csvファイル名_出荷日名
    ship_filename = "出荷日【" & Format(Year(ship_date), "0000") & "年" & Format(Month(ship_date), "00") & "月" & Format(Day(ship_date), "00") & "日" & "】"
    'csvファイル名_現在時刻
    hiduke = Format(Year(Now), "0000") & "_" & Format(Month(Now), "00") & "_" & Format(Day(Now), "00") & "_" & Format(Hour(Now), "00") & "_" & Format(Minute(Now), "00") & "_" & Format(Second(Now), "00")
    
    
    'csvファイル名
    Dim csvFileName As String
    yyyy = Format(Year(ship_date), "0000年")
    mm = Format(Month(ship_date), "00月")
    csvFileName = ActiveWorkbook.Path & "\コープ事前入力csv\" & Customer_name & "\" & yyyy & "\" & mm & "\" & ship_filename & Customer_name & hiduke & ".csv"
    
    'ディレクトリ作成
    Call 事前入力ディレクトリ作成1
    'CSV Open >> Close
    Open csvFileName For Output As #1
    
    Dim i As Long, j As Long
    i = 1
    
    Do While ws.Cells(i, 1).Value <> ""
    
        j = 1
        Do While ws.Cells(i, j + 1).Value <> ""
    
            Print #1, ws.Cells(i, j).Value & ",";
            j = j + 1
    
        Loop
    
        Print #1, ws.Cells(i, j).Value & vbCr;
        i = i + 1
    
    Loop
    
    Close #1

    
    'アクティブ
    Worksheets("ピッキング表").Activate
    Worksheets("ピッキング表").Select
    Range("A1").Select
    
    '非表示
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    
    Application.Calculation = xlCalculationAutomatic    '自動計算
    Worksheets("ピッキング表").Protect '保護

End Sub

Sub 事前入力csv_main()
    'アクティブ
    Worksheets("ピッキング表").Activate
    Worksheets("ピッキング表").Select
    kakunin_ship = Range("D6")
    
    
    
    Dim rc As VbMsgBoxResult
    rc = MsgBox(str(kakunin_ship) + "で事前入力したデータを出力しますか？", vbYesNo + vbQuestion)
    If rc = vbNo Then
        MsgBox "出力を中止します", vbCritical
        Exit Sub
    End If

    Worksheets("ピッキング表").Unprotect   '保護解除
    'A1セルから空白セルまでdowhileで回し、その領域内をcsv出力する
    
    'Application.CalculateFull   '全シート再計算
    
    '非表示
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    
    Call センター数量取得
    
    'アクティブ
    Worksheets("事前出力用").Activate
    Worksheets("事前出力用").Select
    Range("A1").Select
    
    'ワークシート
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets("事前出力用")

    'G2 出荷先
    Dim Customer_name As String
    Customer_name = Range("G2")

    'F1出荷日
    Dim ship_date As Date
    ship_date = Range("D1")

    'csvファイル名_出荷日名
    ship_filename = "出荷日【" & Format(Year(ship_date), "0000") & "年" & Format(Month(ship_date), "00") & "月" & Format(Day(ship_date), "00") & "日" & "】"
    'csvファイル名_現在時刻
    hiduke = Format(Year(Now), "0000") & "_" & Format(Month(Now), "00") & "_" & Format(Day(Now), "00") & "_" & Format(Hour(Now), "00") & "_" & Format(Minute(Now), "00") & "_" & Format(Second(Now), "00")
    
    
    'csvファイル名
    Dim csvFileName As String
    yyyy = Format(Year(ship_date), "0000年")
    mm = Format(Month(ship_date), "00月")
    csvFileName = ActiveWorkbook.Path & "\コープ事前入力csv\" & Customer_name & "\" & yyyy & "\" & mm & "\" & ship_filename & Customer_name & hiduke & ".csv"
    
    'ディレクトリ作成
    Call 事前入力ディレクトリ作成1
    'CSV Open >> Close
    Open csvFileName For Output As #1
    
    Dim i As Long, j As Long
    i = 1
    
    Do While ws.Cells(i, 1).Value <> ""
    
        j = 1
        Do While ws.Cells(i, j + 1).Value <> ""
    
            Print #1, ws.Cells(i, j).Value & ",";
            j = j + 1
    
        Loop
    
        Print #1, ws.Cells(i, j).Value & vbCr;
        i = i + 1
    
    Loop
    
    Close #1

    
    'アクティブ
    Worksheets("ピッキング表").Activate
    Worksheets("ピッキング表").Select
    Range("A1").Select
    
    '非表示
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    
    Application.Calculation = xlCalculationAutomatic    '自動計算
    Worksheets("ピッキング表").Protect '保護
    MsgBox "出力完了しました。"

End Sub

Sub センター数量取得()
    'アクティブ
    Worksheets("ピッキング表").Activate
    Worksheets("ピッキング表").Select
    ship_date = Range("D6")

    Dim center_num As Variant
    ReDim center_num(220, 3) As Variant '取得した行数で2次元配列の再定義
    
    'C列変数宣言
    Dim C_column() As Variant
    ReDim C_column(31, 3) As Variant
    C_column = Range(Cells(8, 3), Cells(38, 3))
    
    num = 0
    For i = 1 To UBound(C_column)
        If C_column(i, 1) = Empty Then
            Exit For
        End If
        
        For c = 1 To 6
'            If c = 5 Then '塩尻の数式
'                   '数式が入っているので飛ばす
'            Else
                If C_column(i, 1) = Empty Or Cells(i + 7, c + 5) = Empty Or Left(C_column(i, 1), 1) <> "Ｉ" Then
                    'デバッグ用
                    If Left(C_column(i, 1), 1) <> "Ｉ" Then
                        Debug.Print C_column(i, 1)
                    End If
                Else
                    center_num(num, 0) = C_column(i, 1)
                    center_num(num, 1) = Cells(i + 7, c + 5)
                    center_num(num, 2) = Cells(6, c + 5)
                    center_num(num, 3) = ship_date
                    num = num + 1
                End If
'            End If
        Next c
    Next i
    
    
    'アクティブ
    Worksheets("事前出力用").Activate
    Worksheets("事前出力用").Select
    Range("A1").Select
    Range(Cells(1, 1), Cells(221, 4)) = center_num

End Sub

Sub 事前入力ディレクトリ作成1()
    Dim root As String
    Dim yyyy As String
    Dim mm As String
    Dim dd As String
    'F1出荷日
    Dim ship_date As Date
    ship_date = Range("D1")

    ' root = ActiveWorkbook.Path & "\csv"
    root = ActiveWorkbook.Path
    yyyy = Format(Year(ship_date), "0000年")
    mm = Format(Month(ship_date), "00月")
    ' dd = Format(Day(ship_date), "00日")

    'G2 出荷先
    Dim Customer_name As String
    Customer_name = Range("G2")

    Dim rtn As Long
    rtn = 事前入力ディレクトリ作成2(root, "コープ事前入力csv", Customer_name, yyyy, mm)
'    Select Case rtn
'        Case 0
'            MsgBox "フォルダを作成しました。"
'        Case 1
'            MsgBox "フォルダは存在します。"
'        Case Else
'            MsgBox "フォルダの作成に失敗しました。"
'    End Select
    
End Sub

Function 事前入力ディレクトリ作成2(ParamArray arg()) As Long
    On Error GoTo ErrExit
    If Dir(Join(arg, "\"), vbDirectory) <> "" Then
        CreateDirectory = 1
        Exit Function
    End If
  
    Dim ary As Variant
    Dim i As Long
    For i = LBound(arg) To UBound(arg)
        ary = arg
        ReDim Preserve ary(i)
        If Dir(Join(ary, "\"), vbDirectory) = "" Then
            MkDir Join(ary, "\")
        End If
    Next
  
    CreateDirectory = 0
    Exit Function
  
ErrExit:
    CreateDirectory = 9
End Function



