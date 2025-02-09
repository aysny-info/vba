
Sub コープデリリセット()
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