Sub ExtractLeadingDecimalNumbersToBE()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim inArr As Variant, outArr() As Variant
    Dim i As Long, j As Long
    Dim cellValue As String
    Dim extracted As String
    Dim currentChar As String
    Dim decimalFound As Boolean
    Dim dblValue As Double

    Set ws = ThisWorkbook.Sheets("原料展開")
    
    ' Q列の最終行を取得（Q3セル以降が対象の場合）
    lastRow = ws.Cells(ws.Rows.Count, "Q").End(xlUp).Row
    If lastRow < 3 Then Exit Sub ' Q列にデータがなければ終了
    
    ' Q列（Q3～最終行）の値を2次元配列として取得
    inArr = ws.Range("Q3:Q" & lastRow).Value
    
    ' 出力用の配列を同じ行数、1列分用意
    ReDim outArr(1 To UBound(inArr, 1), 1 To 1)
    
    ' 各セルごとに先頭連続の数字（小数点を含む）を抽出
    For i = 1 To UBound(inArr, 1)
        ' 全角数字や全角小数点も半角に変換して処理
        cellValue = StrConv(CStr(inArr(i, 1)), vbNarrow)
        extracted = ""
        decimalFound = False
        
        ' 1文字ずつチェック（先頭連続の数字／小数点）
        For j = 1 To Len(cellValue)
            currentChar = Mid(cellValue, j, 1)
            If currentChar Like "[0-9]" Then
                extracted = extracted & currentChar
            ElseIf currentChar = "." Then
                ' 小数点は、既に出現していないかつ先頭に数字がある場合のみ許容
                If Not decimalFound And extracted <> "" Then
                    extracted = extracted & currentChar
                    decimalFound = True
                Else
                    Exit For
                End If
            Else
                Exit For
            End If
        Next j
        
        ' 抽出結果がある場合、末尾が小数点だけの場合は除去してから数値変換を試みる
        If extracted <> "" Then
            If Right(extracted, 1) = "." Then
                extracted = Left(extracted, Len(extracted) - 1)
            End If
            
            ' 数値変換可能なら数値として、エラーなら文字列のまま
            On Error Resume Next
            dblValue = CDbl(extracted)
            If Err.Number = 0 Then
                outArr(i, 1) = dblValue
            Else
                outArr(i, 1) = extracted
            End If
            On Error GoTo 0
        Else
            outArr(i, 1) = ""
        End If
    Next i
    
    ' BE列（BE3～BE最終行）に一括で結果を貼り付け
    ws.Range("BE3:BE" & lastRow).Value = outArr
End Sub
