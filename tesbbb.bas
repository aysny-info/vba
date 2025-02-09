Sub one_search_kai()
'    Application.ScreenUpdating = False                 '画面停止
'    Application.Calculation = xlCalculationManual     '手動計算
'    ActiveSheet.Unprotect      '保護解除

    ThisWorkbook.Activate

    date_D1 = Worksheets("事前入力").Range("D1") '日付
    Dim paste_one_aduke As Variant  '貼り付けデータ
    Dim paste_one_modoshi As Variant  '貼り付けデータ
    Dim LastRow As Long '最終行取得
    Dim LastCol As Long '最終列取得
    
    '*****************コープ取得************************************
    Worksheets("コープ計画数").Activate
    
    LastRow = Cells(Rows.Count, 2).End(xlUp).Row
    
    LastCol = Cells(1, 1).End(xlToRight).Column + 5

    cope_data = Range(Cells(1, 1), Cells(LastRow, LastCol))    '「コープ計画数」シートのデータ

    'cope_dataから絞る用の変数を作成 5日分のデータを格納
    Dim cope_data_5date() As Variant
    ReDim cope_data_5date(1 To LastRow, 1 To LastCol)

    code5_row = 1
    ' cope_dataをdate_D1から4日後までのデータのみに絞る
    For i = 1 To LastRow
        If cope_data(i, 2) >= date_D1 And cope_data(i, 2) <= date_D1 + 4 Then
            For j = 1 To LastCol
                cope_data_5date(code5_row, j) = cope_data(i, j)
            Next j
            code5_row = code5_row + 1
        End If
    Next i

    ' cope_data_5dateの重複なし商品コードを作成
    Dim tyo_code() As Variant
    ReDim tyo_code(1 To LastRow, 1 To 1)

    ' 重複なし商品コードを抽出
    Dim code As Variant
    Dim codeIndex As Long
    codeIndex = 1

    For i = 1 To UBound(cope_data_5date, 1)
        code = cope_data_5date(i, 1) ' 商品コードを仮定
        If Not IsInArray(code, tyo_code) Then
            tyo_code(codeIndex, 1) = code
            codeIndex = codeIndex + 1
        End If
    Next i
    
    Debug.Print 1

    ' cope_data_5date_codeを回して５日分のデータを作成
    Dim paste_data() As Variant
    ReDim paste_data(1 To LastRow, 1 To LastCol)
    p = 1
    arihantei = False

    ' paste_dataのヘッダーを作成 No	商品コード	日付	計画数	商品名
    paste_data(1, 1) = "No"
    paste_data(1, 2) = "商品コード"
    paste_data(1, 3) = "日付"
    paste_data(1, 4) = "計画数"
    paste_data(1, 5) = "商品名"

    ' 5日分のデータを作成　５日で回す
    For d = 0 To 4
        For i = 1 To UBound(tyo_code, 1)
            If Not tyo_code(i, 1) = Empty Then
                For j = 1 To UBound(cope_data_5date, 1)
                    If arihantei = False Then
                        ' 日付と商品コードが一致したら
                        If tyo_code(i, 1) = cope_data_5date(j, 1) And cope_data_5date(j, 2) = date_D1 + d Then
                            paste_data(p, 1) = i
                            For k = 1 To LastCol
                                paste_data(p, k + 1) = cope_data_5date(j, k)
                            Next k
                            p = p + 1
                            arihantei = True
                        End If

                    End If
                     ' 最後の行まできたら
                    If j = UBound(cope_data_5date, 1) Then
                        ' 5日分のデータがなかったら
                        If arihantei = False Then
                            For k = 1 To LastCol
                                paste_data(p, 1) = i
                                paste_data(p, 2) = tyo_code(i, 1)
                                paste_data(p, 3) = date_D1 + d
                                paste_data(p, 4) = 0
                                paste_data(p, 5) = 0
                            Next k

                            p = p + 1
                        End If

                        arihantei = False
                    End If
                Next j
            End If
        Next i
    Next d
        
    'クリアして貼り付け
    Worksheets("コープ形成").Activate
    ActiveSheet.Unprotect      '保護解除
    Worksheets("コープ形成").Cells.ClearContents
    Range(Cells(1, 1), Cells(UBound(paste_data, 1), UBound(paste_data, 2))) = paste_data

    
End Sub