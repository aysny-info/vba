Sub csv_main_2()

    'アクティブ
    ThisWorkbook.Activate
    
    Application.ScreenUpdating = False                 '画面停止
    Application.Calculation = xlCalculationManual     '手動計算
    Call 保護_解除全シート

    ' '日付取得
    ' date_J3 = Worksheets("棚卸明細表").Range("J3") '日付

    ' 'ファイル取得
    ' temp_filename = csvファイル名探索(date_J3)
    
    ' If temp_filename = "" Then
    '     Debug.Print "J3セルの日付の月の原料マスターcsvが見つかりませんでしたので、現在の原料マスターを使用します。"
    '     csvFilePath_genryouM = "\\Afnewt320-kyoyu\社内共有\AFSKS\生産管理\csv\原料マスター_原料マスターシート.csv"
    ' Else
    '     Debug.Print "該当月の原料マスターcsvがあったので読み込みます。" & temp_filename
    '     csvFilePath_genryouM = "\\Afnewt320-kyoyu\社内共有\AFSKS\生産管理\csv\原料マスター履歴\" & temp_filename
    ' End If

    csvFilePath_genryouM = "\\Afnewt320-kyoyu\社内共有\AFSKS\生産管理\csv\原料マスター_原料マスターシート.csv"

    'csv貼り付け_原料M
    sh_name = "原料M"
    shiji_data = getCSV_utf8_2(sh_name, csvFilePath_genryouM)

    Worksheets("棚卸明細表").Activate
    
    Application.Calculation = xlCalculationAutomatic       '手動計算
    Call 保護_全シート
    Application.ScreenUpdating = True                 '画面停止
End Sub

Function getCSV_utf8_2(sh_name As Variant, csvFilePath As Variant) As Variant
    
    'Dim ws As Worksheet
    'Set ws = ThisWorkbook.Worksheets(1)
    
    'Dim strPath As String
    'strPath = "\\Afnewt320-kyoyu\社内共有\AFSKS\在庫表\csv\指示数\【指示数】在庫_ダンボール_2022.3.xlsm.csv"
    
    Dim i As Long, j As Long
    Dim strLine As String
    Dim arrLine As Variant 'カンマでsplitして格納
    
    'D列変数宣言
    Dim paste_data() As Variant
    
    'ADODB.Streamオブジェクトを生成
    Dim adoSt As Object
    Set adoSt = CreateObject("ADODB.Stream")
    
    'シートクリア
    ThisWorkbook.Activate
    Worksheets(sh_name).Activate
    Worksheets(sh_name).Cells.ClearContents
    max_n = 0
    i = 1
    
        With adoSt
            .Charset = "UTF-8"        'Streamで扱う文字コートをutf-8に設定
            .Open                             'Streamをオープン
            .LoadFromFile (csvFilePath) 'ファイルからStreamにデータを読み込む
            
            Do Until .EOS           'Streamの末尾まで繰り返す
                strLine = .ReadText(adReadLine) 'Streamから1行取り込み
                arrLine = Split(Replace(replaceColon_2(strLine), """", ""), ":") 'strLineをカンマで区切りarrLineに格納

                max_n = max_n + 1
            Loop

            .Close
        End With

    '格納する２次元配列サイズ設定
    ReDim paste_data(1 To max_n, 1 To 200) '(行,列)
    
    csv_column_name = 1 'カラム名を１行目に追加
    
      csv_row_num = 1
      
      With adoSt
          .Charset = "UTF-8"        'Streamで扱う文字コートをutf-8に設定
          .Open                             'Streamをオープン
          .LoadFromFile (csvFilePath) 'ファイルからStreamにデータを読み込む
          
          Do Until .EOS           'Streamの末尾まで繰り返す
              
                  strLine = .ReadText(adReadLine) 'Streamから1行取り込み
                  arrLine = Split(Replace(replaceColon_2(strLine), """", ""), ":") 'strLineをカンマで区切りarrLineに格納
                  
                  If csv_column_name = 1 Then 'カラム名を１行目に追加
                      For j = 0 To UBound(arrLine)
                          paste_data(i, j + 1) = arrLine(j)
                      Next j
                      i = i + 1
                      csv_column_name = 2
                  ElseIf csv_row_num <> 1 Then 'データの部分を追加
                      For j = 0 To UBound(arrLine)
                          paste_data(i, j + 1) = arrLine(j)
                      Next j
                      i = i + 1
                  End If
                  
              csv_row_num = csv_row_num + 1
          Loop
      
          .Close
      End With
      
    Range(Cells(1, 1), Cells(max_n, 200)) = paste_data

    getCSV_utf8_2 = paste_data

End Function

'受け取った文字列のカンマをコロンに置き換える
'ダブルクォーテーションで囲まれているカンマは置き換えない
Function replaceColon_2(ByVal str As String) As String
    
    Dim strTemp As String
    Dim quotCount As Long
    
    Dim l As Long
    For l = 1 To Len(str)  'strの長さだけ繰り返す
    
        strTemp = Mid(str, l, 1) 'strから現在の1文字を切り出す
    
        If strTemp = """" Then   'strTempがダブルクォーテーションなら
    
            quotCount = quotCount + 1   'ダブルクォーテーションのカウントを1増やす
    
        ElseIf strTemp = "," Then   'strTempがカンマなら
    
            If quotCount Mod 2 = 0 Then   'quotCountが2の倍数なら
    
                str = Left(str, l - 1) & ":" & Right(str, Len(str) - l)   '現在の1文字をコロンに置き換える
    
            End If
    
        End If
    
    Next l
    
    replaceColon_2 = str

End Function


Function csvファイル名探索(date_J3 As Variant) As Variant
    csvFilePath = "\\Afnewt320-kyoyu\社内共有\AFSKS\生産管理\csv\原料マスター履歴"
    
    'ディレクトリ存在チェック
    If Dir(csvFilePath, vbDirectory) = "" Then
        MsgBox "csvディレクトリが存在しません。" & vbCrLf & csvFilePath
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
        MsgBox "csvファイルが空です。" & vbCrLf & csvFilePath
        End
    End If
    
    ReDim filename(s - 1, 4) '【0】ファイル名.csv【1】作成日【2】年(ファイル名)【3】月0埋め(ファイル名)【4】月0埋め(ファイル名)
    
    With CreateObject("Scripting.FileSystemObject")
        For Each f In .GetFolder(csvFilePath).Files
            filename(cnt, 0) = f.Name
            filename(cnt, 1) = f.DateLastModified   '作成日はf.DateCreated 更新日はf.DateLastModified  アクセス日はf.DateLastAccessed
            cnt = cnt + 1
        Next f
    End With
    
    ' filename(i,1)で、条件が指定された月のfilenameを抽出

    Dim selectedFiles() As String
    Dim newestFile As String
    Dim newestDate As Date
    Dim fileDate As Date

    ' selectedFiles()でインデックス範囲外にアクセスするのを防ぐために、初期化
    ReDim selectedFiles(0)

    For i = 0 To UBound(filename, 1)
        ' filename(cnt, 1)の月とdate_J3の年と月が一致したら、selectedFilesに追加
        If Year(filename(i, 1)) = Year(date_J3) And Month(filename(i, 1)) = Month(date_J3) Then
            ReDim Preserve selectedFiles(UBound(selectedFiles) + 1)
            selectedFiles(UBound(selectedFiles)) = filename(i, 0)
        End If
    Next i

    ' selectedFilesの中から最新のファイルを取得
    For i = 1 To UBound(selectedFiles)
        fileDate = FileDateTime(csvFilePath & "\" & selectedFiles(i))
        If fileDate > newestDate Then
            newestDate = fileDate
            newestFile = selectedFiles(i)
        End If
    Next i

    csvファイル名探索 = newestFile
    
End Function


