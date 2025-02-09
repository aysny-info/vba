Sub 販売計画集計表ボタン()
    Dim rc As VbMsgBoxResult
    rc = MsgBox("販売計画集計表へ反映しますか？", vbYesNo + vbQuestion)
    If rc = vbNo Then
        MsgBox "中止します", vbCritical
        Exit Sub
    End If
    
    Call 販売計画集計表入力main
End Sub

Sub 販売計画集計表入力main()
    han_link = "\\AFnewT320-kyoyu\社内共有\個人フォルダ\笠間\販売計画自動更新\【販売計画集計表】.xlsm"
    ' han_link = "\\Afnewt320-kyoyu\社内共有\【生産管理】\【システム】\【販売計画集計表】.xlsm"
    
    Application.ScreenUpdating = False                 '画面停止
    'Application.Calculation = xlCalculationManual     '手動計算
    'ActiveSheet.Unprotect      '保護解除
    
    Workbooks.Open filename:=han_link, UpdateLinks:=False, IgnoreReadOnlyRecommended:=True
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
    
    Application.ScreenUpdating = True                 '画面
    Application.Calculation = xlAutomatic               '自動計算
    Application.Calculate   ' 計算
    
    Workbooks("【販売計画集計表】.xlsm").Save
    Workbooks("【販売計画集計表】.xlsm").Close

    MsgBox "販売計画集計表へ入力完了しました。" & vbCrLf & vbCrLf & "【出荷日】" & ship_date & vbCrLf & vbCrLf & "正しく入力されているか要確認して下さい。"
    
End Sub

Function csv読み込み(Customer_name As Variant) As Variant
  Dim file As String, max_n As Long
  Dim buf As String, tmp As Variant, ary() As Variant
  Dim i As Long, n As Long, val As Long
  max_n = 0
  
  '準備
  'file = "C:\test.csv" 'ファイル指定
  file = csvファイル名探索(Customer_name)
  
'  max_n = CreateObject("Scripting.FileSystemObject").OpenTextFile(file, 8).Line 'ファイルの行数取得
  
  Open file For Input As #1 'CSVファイルを開く
        Do Until EOF(1)
            Line Input #1, buf
            max_n = max_n + 1
        Loop
  Close #1 'CSVファイルを閉じる
  
  ReDim ary(max_n - 1, 2) As Variant '取得した行数で2次元配列の再定義
    
  Open file For Input As #1 'CSVファイルを開く
      Do Until EOF(1) '最終行までループ
      Line Input #1, buf '読み込んだデータを1行ずつみていく
      tmp = Split(buf, ",") 'カンマで分割
      For i = 0 To UBound(tmp) '項目数ぶんループ
        ary(n, i) = tmp(i) '分割した内容を配列の項目へ入れる（0→ID, 1→名称, 2→値）
      Next i
      n = n + 1 '配列の次の行へ
    Loop
  Close #1 'CSVファイルを閉じる
  
    csv読み込み = ary
'
'  For i = 1 To UBound(ary)
'    Debug.Print ary(i, 0)
'  Next
End Function

Function csvファイル名探索(Customer_name As Variant) As Variant

    ship_date = Sheets("ピッキング表").Range("D6")  '出荷日
    
    'パスの検索
    sFileFullPath = ThisWorkbook.Path
    For i = Len(sFileFullPath) To 0 Step -1
        If InStr(i, sFileFullPath, "\") > 0 Then
            '現在のフォルダ名を取得
            sFolderName = Mid(sFileFullPath, InStr(i, sFileFullPath, "\") + 1)
            '1つ上の階層のフォルダのまでのフルパスを取得
            sParentFolderPath = Mid(sFileFullPath, 1, InStr(1, sFileFullPath, sFolderName) - 2)
            Exit For
        End If
    Next
    csvFilePath = sParentFolderPath & "\ピッキング表\csv\" & Customer_name & "\" & Year(ship_date) & "年\" & Right("0" & Month(ship_date), 2) & "月"
    
    'ディレクトリ存在チェック
    If Dir(csvFilePath, vbDirectory) = "" Then
        MsgBox "csvディレクトリが存在しません。" & vbCrLf & csvFilePath & vbCrLf & vbCrLf & "【出荷日】" & ship_date & vbCrLf & "【出荷先】" & Customer_name & vbCrLf & "①ピッキング表の印刷を行っていない可能性があります。" & vbCrLf & "②生産表A1セルの出荷日が間違えている可能性があります。"
        End
    Else
        Debug.Print "ディレクトリが存在します。"
    End If
    
    Dim f As Object, cnt As Long
    Dim filename() As Variant    '【0】ファイル名.csv【1】作成日【2】年(ファイル名)【3】月(ファイル名)

    cnt = 0
    s = 0
    '２次元配列の格納する数を求める
    With CreateObject("Scripting.FileSystemObject")
        For Each f In .GetFolder(csvFilePath).Files
            s = s + 1
        Next f
    End With
    
    'csv存在チェック
    If s = 0 Then
        MsgBox "csvファイルが空です。" & vbCrLf & csvFilePath & vbCrLf & vbCrLf & "【出荷日】" & ship_date & vbCrLf & "【出荷先】" & Customer_name & vbCrLf & "①ピッキング表の印刷を行っていない可能性があります。" & vbCrLf & "②生産表A1セルの出荷日が間違えている可能性があります。"
        End
    End If
    
    ReDim filename(s - 1, 4) '【0】ファイル名.csv【1】作成日【2】年(ファイル名)【3】月0埋め(ファイル名)【4】月0埋め(ファイル名)
    
    With CreateObject("Scripting.FileSystemObject")
        For Each f In .GetFolder(csvFilePath).Files
            filename(cnt, 0) = f.Name
            filename(cnt, 1) = f.DateLastModified   '作成日はf.DateCreated 更新日はf.DateLastModified  アクセス日はf.DateLastAccessed
            filename(cnt, 2) = Mid(filename(cnt, 0), 5, 4)  '年 ファイル名
            filename(cnt, 3) = Mid(filename(cnt, 0), 10, 2) '月 ファイル名
            filename(cnt, 4) = Mid(filename(cnt, 0), 13, 2) '日 ファイル名
            cnt = cnt + 1
        Next f
    End With
    
    Dim Max As Integer
    Max = 9999 '初期値を設定 下記のfor文で9999のままなら、csvデータに該当の出荷日がなかったということになる。
    For i = 0 To UBound(filename)
        If str(Year(ship_date)) & str(Right("0" & Month(ship_date), 2)) & str(Right("0" & Day(ship_date), 2)) = str(filename(i, 2)) & str(filename(i, 3)) & str(filename(i, 4)) Then
            If Max = 9999 Then
                Max = i
            End If
            If filename(i, 1) > filename(Max, 1) Then
                Max = i
            End If
        End If
    Next i
    
    'csv存在チェック
    If Max = 9999 Then
        MsgBox "該当の出荷日のcsvファイルがありません。" & vbCrLf & csvFilePath & vbCrLf & vbCrLf & "【出荷日】" & ship_date & vbCrLf & "【出荷先】" & Customer_name & vbCrLf & "①ピッキング表の印刷を行っていない可能性があります。" & vbCrLf & "②生産表A1セルの出荷日が間違えている可能性があります。"
        End
    End If
    
    csvファイル名探索 = csvFilePath & "\" & filename(Max, 0)
    
End Function



