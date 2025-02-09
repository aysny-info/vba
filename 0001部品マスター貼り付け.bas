'」」」」」」」 API declarations: 」」」」」」」
Private Declare Sub keybd_event Lib "user32" _
 (ByVal bVk As Byte, _
 ByVal bScan As Byte, _
 ByVal dwFlags As Long, _
 ByVal dwExtraInfo As Long)

Private Declare Function GetKeyboardState Lib "user32" _
 (pbKeyState As Byte) As Long

Const VK_NUMLOCK = &H90           '「NumLock」キー
Const KEYEVENTF_EXTENDEDKEY = &H1 'キーを押す
Const KEYEVENTF_KEYUP = &H2       'キーを放す
Sub numLockOn()
  Dim NumLockState As Boolean
  Dim keys(0 To 255) As Byte

  GetKeyboardState keys(0)
  NumLockState = keys(VK_NUMLOCK)

'「NumLock」キーがオフの場合はオンにする。
  If NumLockState <> True Then
    'キーを押す
    keybd_event VK_NUMLOCK, &H45, KEYEVENTF_EXTENDEDKEY Or 0, 0
    'キーを放す
    keybd_event VK_NUMLOCK, &H45, KEYEVENTF_EXTENDEDKEY Or KEYEVENTF_KEYUP, 0
  End If
End Sub

'Sub データ保存()
''」」」」」」」」」」」」」」」」」」」」」」」」」」」」」」
''」」」」」 ブックを閉じるときに自動実行のためのマクロ
'    Sheets("ピッキング表").Select
''宣言
'    Dim SaveDir1 As String, SaveDir2 As String, _
'        SaveDir3 As String, d As Date, _
'        fn As String, mypath As String
'    d = Range("D6") + 1
''    fn = Worksheets("データ中継").Range("C2") & "_ピッキング表.xlsm"
'    fn = ThisWorkbook.Name
'    mypath = ActiveWorkbook.Path
'    If d >= Date - 14 Then
'    Else
'    Exit Sub
'    End If
''読み取り判定
'    If ActiveWorkbook.ReadOnly = True Then
'    Exit Sub
'    End If
''メッセージ
''    msg = "保存して終了します。" & vbNewLine & _
''          "(保存後自動でブックを閉じます)" & vbNewLine & _
''          "" & vbNewLine & _
''          "※保存しない場合はキャンセルしてください。"
''    kesu = MsgBox(msg, 1, "自動保存")
'    kesu = 1
'    If kesu = 2 Then
'    Exit Sub
'    End If
''上書き保存
'    Application.ScreenUpdating = False
'    Application.DisplayAlerts = False
'    ActiveWorkbook.Save
'    Application.DisplayAlerts = True
''フォルダ存在確認･作成 階層①
'    SaveDir1 = "\\AFNEWT320-KYOYU\社内共有\AFSKS\データ保管\" & Format(d, "yyyy年")
'    If Dir(SaveDir1, vbDirectory) = "" Then
'        MkDir SaveDir1
'    End If
''フォルダ存在確認･作成 階層②
'    SaveDir2 = SaveDir1 & "\" & Format(d, "m月")
'    If Dir(SaveDir2, vbDirectory) = "" Then
'        MkDir SaveDir2
'    End If
''フォルダ存在確認･作成 階層③
'    SaveDir3 = SaveDir2 & "\" & Format(d, "m.d")
'    If Dir(SaveDir3, vbDirectory) = "" Then
'       MkDir SaveDir3
'    End If
''ファイルが元かコピーか確認
'    If mypath Like "*AFSKS\ピッキング表*" Then
'    Else
'    Application.ScreenUpdating = True
'    Exit Sub
'    End If
''階層③フォルダに元Bookを名前を付けて保存
'    Application.DisplayAlerts = False
'    If Dir(SaveDir3 & "\" & fn) = "" Then
'    Else
'    Kill SaveDir3 & "\" & fn
'    End If
'    With ActiveWorkbook
'        .SaveAs Filename:=SaveDir3 & "\" & fn, _
'                          FileFormat:=xlOpenXMLWorkbookMacroEnabled
'    End With
'    On Error Resume Next
'    With ActiveWorkbook
'        .SaveAs Filename:=mypath & "\" & fn, _
'                          FileFormat:=xlOpenXMLWorkbookMacroEnabled
'    End With
'    Application.DisplayAlerts = True
'    On Error GoTo 0
'
''」」」」」」」」」」」」」」」」」」」」」」」」」」」」」」
'End Sub

Sub 注文数入力後振分表印刷()

    Application.ScreenUpdating = False

    ActiveSheet.Unprotect
    ActiveWorkbook.UpdateLink Name:=ActiveWorkbook.LinkSources
    Range("BI6").Value = Now
    Worksheets("ピッキング表").Protect DrawingObjects:=True, Contents:=True, Scenarios:=True
    
    Sheets("振分").Select
    ActiveWindow.SelectedSheets.PrintOut Copies:=1, Collate:=True, _
        IgnorePrintAreas:=False
    Sheets("ピッキング表").Select
    
    Application.ScreenUpdating = True
    
    MsgBox "振分表の印刷が完了しました"
    
End Sub
Sub 印刷前処理()

    Application.ScreenUpdating = False

    ActiveSheet.Unprotect
    ActiveWorkbook.UpdateLink Name:=ActiveWorkbook.LinkSources
    Range("BI7").Value = Now
    Worksheets("ピッキング表").Protect DrawingObjects:=True, Contents:=True, Scenarios:=True
    
    Application.ScreenUpdating = True
    
    MsgBox "帳票印刷前処理が完了しました→②帳票印刷処理へ"
    
End Sub
Sub 印刷_加工一覧()
    Worksheets("加工一覧").Activate
    
    With ActiveSheet
        .Range("B6").Select
        If .FilterMode Then .ShowAllData
            .Range("B6").AutoFilter Field:=1, Criteria1:="<>"
            
            'ソート
            ActiveWorkbook.Worksheets("加工一覧").AutoFilter.Sort.SortFields.Add2 Key:=Range _
            ("C4:C75"), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:= _
            xlSortNormal
            With ActiveWorkbook.Worksheets("加工一覧").AutoFilter.Sort
                .Header = xlYes
                .MatchCase = False
                .Orientation = xlTopToBottom
                .SortMethod = xlPinYin
                .Apply
            End With
        
    End With
    
    With ActiveSheet.PageSetup
        .PaperSize = xlPaperA3
        
        .TopMargin = Application.CentimetersToPoints(1)
        .BottomMargin = Application.CentimetersToPoints(1)
        .LeftMargin = Application.CentimetersToPoints(1)
        .RightMargin = Application.CentimetersToPoints(1)
        .HeaderMargin = Application.CentimetersToPoints(0.5)
        .FooterMargin = Application.CentimetersToPoints(0.5)
    End With
    
    ActiveWindow.SelectedSheets.PrintOut Copies:=1
End Sub
Sub ユーコープP印刷()
    '2021.01.31
    Call csv出力.csv_main
    '2023.06.07
    Call 実績csv_main

    '」」」」」」」」」」」」」」」」」」」」」」」」」」」」
    'データエラーチェック
    If Range("BG9") = "OK" Then
        'pass
    Else
        MsgBox "※データ異常あり(印刷をキャンセルします)"
        Exit Sub
    End If

    If Worksheets("ピッキング表").Range("BG9") = "NG" Then
        MsgBox "▲▲▲更新データに異常があります。" & vbNewLine & "ファイルを閉じて更新しなおしてください。"
        Exit Sub
    End If

    n1 = Range("BF14")
    n2 = Range("BF15")
    n3 = Range("BF16")
    n4 = Range("BF17")

    msg = "印刷内容を確認してください。" & _
          vbNewLine & n1 & "   " & _
          vbNewLine & n2 & "   " & _
          vbNewLine & n3 & "   " & _
          vbNewLine & n4 & "   "

    kesu = MsgBox(msg, 1, "帳票発行")

    If kesu = 2 Then
        Exit Sub
    End If

    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    ActiveWorkbook.Save
    Application.DisplayAlerts = True

    If Worksheets("ピッキング表").Range("BG9") = "NG" Then
        MsgBox "▲▲▲更新データに異常があります。" & vbNewLine & _
            "ファイルを閉じて更新しなおしてください。"
        Application.ScreenUpdating = True
        Worksheets("ピッキング表").Protect DrawingObjects:=True, Contents:=True, Scenarios:=True
        Exit Sub
    End If

'印刷処理1
    Sheets("振分").Select
    ActiveWindow.SelectedSheets.PrintOut Copies:=2, Collate:=True, _
        IgnorePrintAreas:=False

    Sheets("ラベル用").Select
    ActiveWindow.SelectedSheets.PrintOut Copies:=1, Collate:=True, _
        IgnorePrintAreas:=False

    Sheets("レシピ用").Select
    ActiveWindow.SelectedSheets.PrintOut Copies:=1, Collate:=True, _
        IgnorePrintAreas:=False

    Sheets("レシピ看板").Select
    ActiveSheet.Range("A1").AutoFilter Field:=1, Criteria1:="<>"
    ActiveWindow.SelectedSheets.PrintOut Copies:=1, Collate:=True, _
        IgnorePrintAreas:=False
    ActiveSheet.Range("A1").AutoFilter Field:=1

    '2023.01.29　非コメントブロック
    Sheets("取引依頼書 (森の里)").Select
    ActiveSheet.Range("A1").AutoFilter Field:=1, Criteria1:="<>"
    ActiveWindow.SelectedSheets.PrintOut Copies:=2, Collate:=True, _
        IgnorePrintAreas:=False
    ActiveSheet.Range("A1").AutoFilter Field:=1

    '2024.05.19 5月5回より沼津センター閉鎖
'    Sheets("取引依頼書 (沼津)").Select
'    ActiveSheet.Range("A1").AutoFilter Field:=1, Criteria1:="<>"
'    ActiveWindow.SelectedSheets.PrintOut Copies:=2, Collate:=True, _
'        IgnorePrintAreas:=False
'    ActiveSheet.Range("A1").AutoFilter Field:=1
    
    '2024.02.10
    Sheets("取引依頼書 (駿河)").Select
    ActiveSheet.Range("A1").AutoFilter Field:=1, Criteria1:="<>"
    ActiveWindow.SelectedSheets.PrintOut Copies:=2, Collate:=True, _
        IgnorePrintAreas:=False
    ActiveSheet.Range("A1").AutoFilter Field:=1

    '印刷回数履歴
    Worksheets("ピッキング表").Range("BJ14").Formula = Worksheets("ピッキング表").Range("BJ14") + 1

    If Worksheets("ピッキング表").Range("BG9") = "NG" Then
        MsgBox "▲▲▲更新データに異常があります。" & vbNewLine & _
            "ファイルを閉じて更新しなおしてください。"
        Application.ScreenUpdating = True
        Worksheets("ピッキング表").Protect DrawingObjects:=True, Contents:=True, Scenarios:=True
        Worksheets("ピッキング表").Select
        Exit Sub
    End If

    '2023.07.02 来週～初日だけ検査員1部（全部で2部）、その他は1部
    Call フォント調整_規格
    With Worksheets("チェックシート")
        .Visible = True
        .Range("A1").AutoFilter Field:=1, Criteria1:="<>"

        If Worksheets("ピッキング表").Range("T6").Value = "土" Then
            .PrintOut Copies:=2, Collate:=True, IgnorePrintAreas:=False
        Else
            .PrintOut Copies:=1, Collate:=True, IgnorePrintAreas:=False
        End If

        .Range("A1").AutoFilter Field:=1

    End With


    '印刷回数履歴
    Worksheets("ピッキング表").Range("BJ15").Formula = Worksheets("ピッキング表").Range("BJ15") + 1

    If Worksheets("ピッキング表").Range("BG9") = "NG" Then
    MsgBox "▲▲▲更新データに異常があります。" & vbNewLine & _
           "ファイルを閉じて更新しなおしてください。"
    Application.ScreenUpdating = True
    Worksheets("ピッキング表").Protect DrawingObjects:=True, Contents:=True, Scenarios:=True
    Worksheets("ピッキング表").Select
    Exit Sub
    End If

    '印刷処理3
    Sheets("ローラー掛け").Select
    ActiveSheet.Range("A1").AutoFilter Field:=1, Criteria1:="<>"
    ActiveWindow.SelectedSheets.PrintOut Copies:=1, Collate:=True, _
        IgnorePrintAreas:=False
    ActiveSheet.Range("A1").AutoFilter Field:=1

    Sheets("ロットメモ").Select
    ActiveSheet.Range("A1").AutoFilter Field:=1, Criteria1:="<>"
    ActiveWindow.SelectedSheets.PrintOut Copies:=1, Collate:=True, _
        IgnorePrintAreas:=False
    ActiveSheet.Range("A1").AutoFilter Field:=1

    Sheets("作業順番表").Select
    ActiveSheet.Range("I6").AutoFilter Field:=9, Criteria1:="<>"
    ActiveWindow.SelectedSheets.PrintOut Copies:=1, Collate:=True, _
        IgnorePrintAreas:=False
    ActiveSheet.Range("I6").AutoFilter Field:=9

    Sheets("看板").Select
    ActiveSheet.Range("A1").AutoFilter Field:=1, Criteria1:="<>"
    ActiveWindow.SelectedSheets.PrintOut Copies:=2, Collate:=True, _
        IgnorePrintAreas:=False
    ActiveSheet.Range("A1").AutoFilter Field:=1

    Sheets("看板2").Select
    ActiveSheet.Range("A1").AutoFilter Field:=1, Criteria1:="<>"
    ActiveWindow.SelectedSheets.PrintOut Copies:=1, Collate:=True, _
        IgnorePrintAreas:=False
    ActiveSheet.Range("A1").AutoFilter Field:=1

    Sheets("看板3").Select
    ActiveSheet.Range("A1").AutoFilter Field:=1, Criteria1:="<>"
    ActiveWindow.SelectedSheets.PrintOut Copies:=1, Collate:=True, _
        IgnorePrintAreas:=False
    ActiveSheet.Range("A1").AutoFilter Field:=1

    Sheets("看板4").Select
    ActiveSheet.Range("A1").AutoFilter Field:=1, Criteria1:="<>"
    ActiveWindow.SelectedSheets.PrintOut Copies:=1, Collate:=True, _
        IgnorePrintAreas:=False
    ActiveSheet.Range("A1").AutoFilter Field:=1

    Sheets("看板5").Select
    ActiveSheet.Range("A1").AutoFilter Field:=1, Criteria1:="<>"
    ActiveWindow.SelectedSheets.PrintOut Copies:=1, Collate:=True, _
        IgnorePrintAreas:=False
    ActiveSheet.Range("A1").AutoFilter Field:=1

    Application.Wait Time:=Now + TimeValue("00:00:30")

    '印刷回数履歴
    Worksheets("ピッキング表").Range("BJ16").Formula = Worksheets("ピッキング表").Range("BJ16") + 1

    If Worksheets("ピッキング表").Range("BG9") = "NG" Then
        MsgBox "▲▲▲更新データに異常があります。" & vbNewLine & _
            "ファイルを閉じて更新しなおしてください。"
        Application.ScreenUpdating = True
        Worksheets("ピッキング表").Protect DrawingObjects:=True, Contents:=True, Scenarios:=True
        Worksheets("ピッキング表").Select
        Exit Sub
    End If

    '指示書的印刷処理
    Sheets("払い出し一覧").Select

    ActiveWorkbook.Worksheets("払い出し一覧").AutoFilter.Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("払い出し一覧").AutoFilter.Sort.SortFields.Add Key:=Range("E4"), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:= xlSortNormal
    With ActiveWorkbook.Worksheets("払い出し一覧").AutoFilter.Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With

    ActiveSheet.Range("A1").AutoFilter Field:=1, Criteria1:="<>"
    ActiveWindow.SelectedSheets.PrintOut Copies:=1, Collate:=True, IgnorePrintAreas:=False
    ActiveSheet.Range("A1").AutoFilter Field:=1

    ActiveWorkbook.Worksheets("払い出し一覧").AutoFilter.Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("払い出し一覧").AutoFilter.Sort.SortFields.Add Key:=Range("B4"), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:= xlSortNormal
    With ActiveWorkbook.Worksheets("払い出し一覧").AutoFilter.Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With

    '2022.10.04 追加
    Sheets("払出加工表").Select
    ActiveSheet.Range("A1").AutoFilter Field:=1, Criteria1:="<>"
    ActiveWindow.SelectedSheets.PrintOut Copies:=1, Collate:=True, _
        IgnorePrintAreas:=False
    ActiveSheet.Range("A1").AutoFilter Field:=1

'    Application.Wait Time:=Now + TimeValue("00:00:10")

    Sheets("包装表").Select
    ActiveSheet.Range("A1").AutoFilter Field:=1, Criteria1:="<>"
    ActiveWindow.SelectedSheets.PrintOut Copies:=1, Collate:=True, _
        IgnorePrintAreas:=False
    ActiveSheet.Range("A1").AutoFilter Field:=1

    '2023.11.18移動（A3紙をまとめて印刷）
    With Worksheets("振分(出荷)")
        'PageSetup.Zoom = 15
        .PageSetup.PaperSize = xlPaperA4
    End With
    Sheets("振分(出荷)").Select
    ActiveWindow.SelectedSheets.PrintOut Copies:=4, Collate:=True, _
        IgnorePrintAreas:=False

    With Worksheets("振分(出荷)")
        'PageSetup.Zoom = 22
        .PageSetup.PaperSize = xlPaperA3
    End With
    Sheets("振分(出荷)").Select
    ActiveWindow.SelectedSheets.PrintOut Copies:=1, Collate:=True, _
        IgnorePrintAreas:=False
    
    Sheets("ユー日別").Select
    Worksheets("ユー日別").PageSetup.PaperSize = xlPaperA3
    ActiveSheet.Range("A1").AutoFilter Field:=1, Criteria1:="<>"
    ActiveWindow.SelectedSheets.PrintOut Copies:=1, Collate:=True, _
        IgnorePrintAreas:=False
    ActiveSheet.Range("A1").AutoFilter Field:=1

    '2023.06.23
    Call 印刷_加工一覧

    '印刷回数履歴
    Worksheets("ピッキング表").Range("BJ17").Formula = Worksheets("ピッキング表").Range("BJ17") + 1

    Worksheets("ピッキング表").Select
    ActiveSheet.Protect DrawingObjects:=True, Contents:=True, Scenarios:=True
    Range("F8").Select
    Call データ保管.データ保管
    '2022.07.30
    Call 販売計画集計表入力main
    '2023.05.07
    Workbooks("ユーコープ_ピッキング表.xlsm").Activate
    Call 形成用
    '2023.04.14
    Call 出荷数管理表入力ボタン
    '2023.06.07
'    call 販売計画実績集計表入力main
    
    ThisWorkbook.Activate
    Worksheets("ピッキング表").Activate
    Worksheets("ピッキング表").Select
    Range("A1").Select
    Application.DisplayAlerts = False
    ActiveWorkbook.Save
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    
End Sub

Sub ユーコープ全表示()
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual ' 再計算を手動に設定
    ActiveSheet.Unprotect
    
    ' 全ての行と列を非表示解除
    Cells.EntireRow.Hidden = False
    Cells.EntireColumn.Hidden = False
    
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic ' 再計算を自動に戻す
End Sub
Sub ユーコープ非表示()

    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual ' 再計算を手動に設定
    
    ' 直接行を非表示にする
    Rows("4:5").Hidden = True
    Rows("28:48").Hidden = True
    
    ' 直接列を非表示にする
    Columns("B:B").Hidden = True
    Columns("H:BC").Hidden = True
    Columns("BK:BL").Hidden = True
    
    ' シートを保護する
    ActiveSheet.Protect DrawingObjects:=True, Contents:=True, Scenarios:=True
    
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic ' 再計算を自動に戻す
End Sub
Sub ユーコープリセット()
    msg = "受注数をクリアし、出荷日を変更し、その出荷日の商品一覧を反映させます。"
    kesu = MsgBox(msg, 1, "リセット")
    If kesu = 2 Then
        Exit Sub
    End If
    
    '翌日の日付を取得
    defaultDate = Date + 1

    ' ユーザーに日付を入力させる（デフォルトで計算された日付を表示）
    userDate = InputBox("出荷日を入力してください " & Format(defaultDate, "yyyy/mm/dd aaa "), "出荷日入力", Format(defaultDate, "yyyy/mm/dd"))
    
    ' ユーザーが何も入力しなかった場合、デフォルトの日付を使用
    If IsDate(userDate) Then
        userDate = CDate(userDate)
    Else
        userDate = defaultDate
    End If
    
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual '計算手動

    
    ' シートを保護解除し、データをクリア
    Worksheets("ピッキング表").Unprotect
    Range("F8:L32,BJ14:BJ17,BI6:BI7").Select
    Selection.ClearContents
    ' 出荷日を設定
    Range("D6").Formula = userDate
    Worksheets("ピッキング表").Protect DrawingObjects:=True, Contents:=True, Scenarios:=True, AllowSorting:=True, AllowFiltering:=True

    Call インポート_販売計画csv

    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic '計算自動
    
    MsgBox ("リセット終了しました。")
End Sub



