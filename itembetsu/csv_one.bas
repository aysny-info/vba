Sub ImportCSVDirectly_Fastest_WithClear()
    Dim csvFilePath As String
    Dim ws As Worksheet

    ' CSVファイルのパスを指定
    csvFilePath = "\\AFnewT320-kyoyu\社内共有\個人フォルダ\笠間\item_betsu\2024年アイテム別出荷数量.xlsm.csv"
    
    ' データを貼り付けるシートを設定
    Set ws = ThisWorkbook.Sheets("Sheet1") ' 必要に応じてシート名を変更
    
    ' シートをクリア
    ws.Cells.Clear

    ' Excelの描画更新を無効化
    Application.ScreenUpdating = False

    ' CSVファイルを直接シートにインポート
    With ws.QueryTables.Add(Connection:="TEXT;" & csvFilePath, Destination:=ws.Range("A1"))
        .TextFileParseType = xlDelimited
        .TextFileConsecutiveDelimiter = False
        .TextFileTabDelimiter = False
        .TextFileSemicolonDelimiter = False
        .TextFileCommaDelimiter = True ' カンマ区切りを指定
        .TextFileTextQualifier = xlTextQualifierDoubleQuote ' ダブルクォートを考慮
        .TextFilePlatform = 65001 ' UTF-8エンコード
        .Refresh BackgroundQuery:=False
    End With

    ' 描画更新を元に戻す
    Application.ScreenUpdating = True

    MsgBox "シートのクリアとCSVファイルのインポートが完了しました。"
End Sub

