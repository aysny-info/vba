Sub 集計表の更新()
    ' han_link = "\\AFnewT320-kyoyu\社内共有\個人フォルダ\笠間\販売計画自動更新\【販売計画集計表】.xlsm"
    han_link = "\\Afnewt320-kyoyu\社内共有\【生産管理】\【システム】\【販売計画集計表】.xlsm"
    
    Application.ScreenUpdating = False                 '画面停止
    'Application.Calculation = xlCalculationManual     '手動計算
    'ActiveSheet.Unprotect      '保護解除
    
    Workbooks.Open Filename:=han_link, UpdateLinks:=False, IgnoreReadOnlyRecommended:=True
    Workbooks("【販売計画集計表】.xlsm").Activate
    
    '読み取りかどうか
    Dim wb  As Workbook
    Dim readOnlyFlg
    
    Set wb = ActiveWorkbook
    readOnlyFlg = wb.ReadOnly
    
    If readOnlyFlg = True Then
        Debug.Print "読み取り専用です"
        MsgBox "販売計画集計表は読み取り専用です。どなたか使用していますので再チャレンジして下さい。"
        Workbooks("【販売計画集計表】.xlsm").Close SaveChanges:=False

        Application.ScreenUpdating = True                 '画面
        Application.Calculation = xlAutomatic               '自動計算
        End
    Else
        Debug.Print "読み取り専用ではありません"
    End If
    
    Application.ScreenUpdating = True                   '画面
    Application.Calculation = xlAutomatic               '自動計算
    Application.Calculate                               ' 計算
    
    Workbooks("【販売計画集計表】.xlsm").Save
    Workbooks("【販売計画集計表】.xlsm").Close

    MsgBox ("販売計画集計表の更新が完了しました。")

End Sub

