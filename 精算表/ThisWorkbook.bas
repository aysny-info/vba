Private Sub Workbook_Open()
    On Error GoTo ErrorHandler  ' エラー発生時のジャンプ先を設定

    Application.Calculation = xlManual  ' 手動計算に切替
    Application.ScreenUpdating = False                 '画面停止

    Call 保護.全保護解除
    Call 開いた日時
    ' 参照設定の追加
    Call Microsoft_Scripting_Runtime
    Call microsoft_visual_basic_for_applications_extensibility_5_3
    Call microsoft_activex_data_objects_2_8_library

    Call csv読み込み.マスター読み込み
    Call 保護.複数保護
    Call デフォルト位置

CleanUp:
    Application.Calculation = xlAutomatic   ' 自動計算に戻す
    '再計算する
    Application.CalculateFull                      '全計算
    Call BE列計算
    Application.ScreenUpdating = True                 '画面停止
    Exit Sub

ErrorHandler:
    ' エラー発生時の処理（必要に応じてログ出力なども追加してください）
    MsgBox "エラーが発生しました:" & vbCrLf & Err.Description, vbExclamation, "エラー"
    Resume CleanUp
End Sub

Sub 開いた日時()
    Worksheets("マスター更新日時").Range("B3").Value = Now
End Sub

Sub デフォルト位置()
    'アクティブ
    Worksheets("受注入力").Activate
    Worksheets("受注入力").Select
    Range("A9").Select
End Sub

Sub aa()
    '参照設定の追加
    Call Microsoft_Scripting_Runtime
    Call microsoft_visual_basic_for_applications_extensibility_5_3
    Call microsoft_activex_data_objects_2_8_library
End Sub
