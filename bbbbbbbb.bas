Option Explicit

'★ 数式を残したい列(アルファベット)を指定
Private Const COLS_TO_KEEP_FORMULA As String = "P,Q,R,Y,AB,AE"

Public Sub 原料展開_数式を値に_改修版_Simple()
    Dim wsDst As Worksheet
    Dim 最終行 As Long
    Dim 消す行 As Long
    
    '【1】シートと最終行を設定
    Set wsDst = ThisWorkbook.Sheets("test")
    最終行 = 10          ' ★本来は UsedRange や最終行検索で取得してください
    消す行 = 7000        ' ★大きめの行番号を指定して一括クリア
    
    '----------------------------------------------------------------------
    ' 1) 画面更新停止 & 再計算モードを手動化
    '----------------------------------------------------------------------
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    
    '----------------------------------------------------------------------
    ' 2) 数式領域クリア
    '----------------------------------------------------------------------
    wsDst.Range("A4:BD" & 消す行).ClearContents
    
    '----------------------------------------------------------------------
    ' 3) オートフィル (例: A3:BD3 → A3:BD 最終行)
    '----------------------------------------------------------------------
    wsDst.Range("A3:BD3").AutoFill Destination:=wsDst.Range("A3:BD" & 最終行)
    
    '----------------------------------------------------------------------
    ' 4) 手動モードで一度強制計算 (新たに貼り付けた数式を計算させる)
    '----------------------------------------------------------------------
    Application.Calculate
    
    '----------------------------------------------------------------------
    ' 5) 数式を残したい列を配列に
    '----------------------------------------------------------------------
    Dim skipCols As Variant
    skipCols = Split(COLS_TO_KEEP_FORMULA, ",")  ' 例: {"P","Q","R","Y","AB","AE"}
    
    '----------------------------------------------------------------------
    ' 6) 1列ずつ「値」に変換 (スキップ列は除外)
    '----------------------------------------------------------------------
    Dim colNum As Long
    Dim colLetter As String
    
    Dim firstRow As Long
    firstRow = 4
    
    Dim lastRow As Long
    lastRow = 最終行
    
    Dim rng As Range
    
    Dim t As Single
    t = Timer  ' ★処理時間計測用(必要なければ削除)
    
    For colNum = 1 To 56  ' A列(1)～BD列(56)
        
        ' 列番号→列文字を取得する関数（後述のColumnLetter）を使用
        colLetter = ColumnLetter(colNum)
        
        ' スキップ列に含まれなければ 4行目～最終行を値に変換
        If Not IsInArray(colLetter, skipCols) Then
            Set rng = wsDst.Range(colLetter & firstRow & ":" & colLetter & lastRow)
            rng.Value = rng.Value
        End If
    Next colNum
    
    Debug.Print "値変換に要した時間：" & (Timer - t) & " 秒"  ' ★処理時間を即時出力

    '----------------------------------------------------------------------
    ' 7) 処理終了前に再度計算モードを自動に戻す + 再計算 & 画面更新再開
    '----------------------------------------------------------------------
    Application.Calculation = xlCalculationAutomatic
    Application.Calculate
    Application.ScreenUpdating = True
    
    MsgBox "完了しました。"
End Sub

'--- 列番号(1→A, 2→B, ..., 27→AA etc.) を列文字に変換 ---
Private Function ColumnLetter(ByVal colNum As Long) As String
    ' 例: Cells(1,1).Address(True,True) → "$A$1" → Split → {"","A","1"}
    ColumnLetter = Split(Cells(1, colNum).Address(True, True), "$")(1)
End Function

'--- 配列内に文字列が含まれるかチェックする関数 ---
Private Function IsInArray(ByVal val As String, ByVal arr As Variant) As Boolean
    Dim element As Variant
    For Each element In arr
        If StrComp(Trim(element), Trim(val), vbTextCompare) = 0 Then
            IsInArray = True
            Exit Function
        End If
    Next element
End Function
