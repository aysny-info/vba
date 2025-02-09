Private Sub Workbook_Open()
    Application.Calculation = xlAutomatic
    '参照設定の追加
    Call Microsoft_Scripting_Runtime
    Call microsoft_visual_basic_for_applications_extensibility_5_3
    Call microsoft_activex_data_objects_2_8_library

    ThisWorkbook.Activate
    first_sheet = ActiveSheet.Name

    On Error GoTo ErrorHandler '例外処理 2021.07.31
        '****************************   ピッキング表判定 2021.07.31  ****************************
        Dim z As Long, tmp As String
        For z = 1 To Workbooks.Count
            tmp = tmp & Workbooks(z).Name & vbCrLf
        Next
        
        ' チェックしたいキーワードを配列に格納
        Keywords = Array("ピッキング表", "AF出荷商品一覧", "コープデリフローズン.xlsm", "ピースデリ.xlsm", "原信.xlsm","フレッセイ.xlsm","コープ.xlsm","IY.xlsm","ユーコープ.xlsm","ヨーク.xlsm","CGC.xlsm","SMコープデリ.xlsm","IY惣菜.xlsm","セイミヤ.xlsm","東武.xlsm","文化堂.xlsm","津田商店.xlsm","原信(日本アクセス).xlsm","ヤオコー.xlsm")
        isMatch = False
        ' 配列内のキーワードを順番に確認
        For Each keyword In Keywords
            If tmp Like "*" & keyword & "*" Then
                isMatch = True
                Exit For ' 一度マッチしたらループを抜ける
            End If
        Next keyword
        
        ' ****************************    ピッキング表開いてる場合    ****************************
        If isMatch Then
            Debug.Print "ピッキング表や販売計画などが開いているので、メッセージボックスはスルー"
        Else
            ' ****************************    ピッキング表開いていない場合    ****************************
            'csvのインポート 2022.07.30
            Call csv_main
             
            Worksheets("未登録商品").Visible = True
            
            ' first_sheet = "処理日"
            Worksheets("未登録商品").Activate
    
            yet_list = Worksheets("未登録商品").Range(Cells(2, 1), Cells(12, 3))
            Dim msg1 As String
    
            For i = 1 To 10
                If yet_list(i, 1) <> "" Then
                    msg1 = msg1 & yet_list(i, 1) & " 残り" & DateDiff("d", Now, yet_list(i, 2)) & "日 " & yet_list(i, 3) & vbCrLf
                End If
            Next
            Worksheets(first_sheet).Activate
    

            
            MsgBox "未登録商品(10件まで表示)" & vbCrLf & vbCrLf & msg1, vbExclamation
        End If

'例外処理
ErrorHandler:
    'Finally:へ飛ぶ
    Resume Finally
'最終処理
Finally:
     Worksheets(first_sheet).Activate
     Worksheets("未登録商品").Visible = False
End Sub


