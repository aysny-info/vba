Option Explicit

'★ 数式を残したい列(アルファベット)を指定
Private Const COLS_TO_KEEP_FORMULA As String = "P,Q,R,Y,AB,AE"

Public Sub 原料展開_数式を値に_改修版_Simple()
    Dim wsDst As Worksheet
    Dim 最終行 As Long
    Dim 消す行 As Long
    
    '【1】シートと最終行を設定
    Set wsDst = ThisWorkbook.Sheets("test")
    最終行 = 10
    
    消す行 = 7000
    
    
    '【1.5】4行目～最終行までの内容(式含む)を消去
    wsDst.Range("A4:BD" & 消す行).ClearContents
    
    '【2】オートフィル (例: A3:BD3 → A3:BD10)
    wsDst.Range("A3:BD3").AutoFill Destination:=wsDst.Range("A3:BD" & 最終行)
    
    '【3】数式を残したい列を配列に
    Dim skipCols As Variant
    skipCols = Split(COLS_TO_KEEP_FORMULA, ",")  ' → {"P","Q","R","Y"}
    
    '【4】1列ずつ「値」に変換 (スキップ列は除外)
    Dim colNum As Long
    Dim colLetter As String
    
    ' なるべく画面更新や再計算の負荷を下げる
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    
    For colNum = 1 To 56  ' A(1)～BD(56)
        
        ' 列番号 → 列文字(A,B,C,...)
        colLetter = ColumnLetter(colNum)
        
        ' スキップ列に含まれなければ、4行目～最終行を値に変換
        If Not IsInArray(colLetter, skipCols) Then
            With wsDst.Range(colLetter & "4:" & colLetter & 最終行)
                .Value = .Value
            End With
        End If
        
    Next colNum
    
    ' 終了処理
    Application.Calculation = xlCalculationAutomatic
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

