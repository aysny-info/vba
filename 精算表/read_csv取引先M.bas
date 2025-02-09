Sub csv_main_01()
    'アクティブ
    ThisWorkbook.Activate
    
    ' Application.ScreenUpdating = False                 '画面停止
    ' Application.Calculation = xlCalculationManual     '手動計算
    ' ActiveSheet.UnProtect      '保護解除
    
    ' ***************       csv1つ目の読み込み
    csvFilePath_shoM = "\\AFnewT320-kyoyu\社内共有\AFSKS\生産管理\csv\取引先マスター.csv"

    'csv貼り付け
    sh_name = "取引先Mcsv"
    Dim sho_data As Variant
    sho_data = getCSV_utf8_2(sh_name, csvFilePath_shoM)
    
    ' Application.ScreenUpdating = True                 '画面停止
    ' Application.Calculation = xlCalculationAutomatic       '手動計算
End Sub

Function getCSV_utf8_2(sh_name As Variant, csvFilePath As Variant) As Variant
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
    ReDim paste_data(1 To max_n, 1 To 210) '(行,列)
    
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
      
    Range(Cells(1, 1), Cells(max_n, 210)) = paste_data

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


