Sub 計算をVBA内で実施して値貼付()
    Dim wsK As Worksheet, wsY As Worksheet
    Dim rX As Range
    Dim i As Long, j As Long
    Dim lastRowAT As Long
    Dim arrAT As Variant, arrBD As Variant
    Dim arrD As Variant, arrG As Variant
    Dim result As Variant
    Dim currentRow As Long
    Dim minVal As Variant
    Dim found As Boolean
    
    ' シートの設定
    Set wsK = ThisWorkbook.Worksheets("加工量")
    Set wsY = ThisWorkbook.Worksheets("原料展開")
    
    ' 対象範囲（X7:X10）を設定
    Set rX = wsK.Range("X7:600")
    
    ' 加工量シートのD列とG列の必要な部分を取得（ここではX列の行番号に合わせる）
    Dim startRow As Long, endRow As Long
    startRow = rX.Row
    endRow = rX.Row + rX.Rows.Count - 1
    arrD = wsK.Range("D" & startRow & ":D" & endRow).Value
    arrG = wsK.Range("G" & startRow & ":G" & endRow).Value
    
    ' 原料展開シートのAT列とBD列の最終行を特定して配列に読み込む
    lastRowAT = wsY.Cells(wsY.Rows.Count, "AT").End(xlUp).Row
    arrAT = wsY.Range("AT1:AT" & lastRowAT).Value
    arrBD = wsY.Range("BD1:BD" & lastRowAT).Value
    
    ' 結果用の配列を初期化（1行分ずつ）
    ReDim result(1 To rX.Rows.Count, 1 To 1)
    
    ' 各対象行ごとに計算を実施
    For i = 1 To UBound(result, 1)
        currentRow = startRow + i - 1
        
        ' G列が0以下なら結果は "-" とする
        If arrG(i, 1) <= 0 Then
            result(i, 1) = "-"
        Else
            ' G列が正の場合、原料展開シートのAT列が加工量シートのD列と一致する行のBD列からMINを取得
            found = False
            minVal = 0 ' 初期値（仮）
            For j = 1 To UBound(arrAT, 1)
                If arrAT(j, 1) = wsK.Cells(currentRow, "D").Value Then
                    ' 数値かどうかチェック（または "-" でないか）
                    If IsNumeric(arrBD(j, 1)) Then
                        If Not found Then
                            minVal = arrBD(j, 1)
                            found = True
                        Else
                            If arrBD(j, 1) < minVal Then
                                minVal = arrBD(j, 1)
                            End If
                        End If
                    End If
                End If
            Next j
            If found Then
                result(i, 1) = minVal
            Else
                result(i, 1) = "-"
            End If
        End If
    Next i
    
    ' 結果をX列に貼付（値のみ）
    rX.Value = result
    
    MsgBox "処理が完了しました。"
End Sub
