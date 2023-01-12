Sub 数式反映(sh_name As String)
    Application.ScreenUpdating = False              '画面停止
    Application.Calculation = xlCalculationManual   '計算停止
    
    Sheets(sh_name).Activate                        'シートアクティブ

    '---------------------------------------        単品            ---------------------------------------
    'B12～B27にindex matchの数式を入力  部門
    For i = 12 To 27
        Range("B" & i).Formula = "=IFERROR(INDEX(" & sh_name & "_単!$I:$I,MATCH($A" & i & "," & sh_name & "_単!$A:$A,0)),"""")"
    Next i

    'C12～C27にindex matchの数式を入力  商品名
    For i = 12 To 27
        Range("C" & i).Formula = "=IFERROR(INDEX(" & sh_name & "_単!$C:$C,MATCH($A" & i & "," & sh_name & "_単!$A:$A,0)),"""")"
    Next i

    'D12～D27にindex matchの数式を入力  入数
    For i = 12 To 27
        Range("D" & i).Formula = "=IFERROR(INDEX(" & sh_name & "_単!$E:$E,MATCH($A" & i & "," & sh_name & "_単!$A:$A,0)),"""")"
    Next i

    'E12～E27にindex matchの数式を入力  数量
    For i = 12 To 27
        Range("E" & i).Formula = "=IFERROR(INDEX(" & sh_name & "_単!$F:$F,MATCH($A" & i & "," & sh_name & "_単!$A:$A,0)),"""")"
    Next i
    
    'G12～G18にindex matchの数式を入力  部門
    For i = 12 To 18
        Range("G" & i).Formula = "=IFERROR(INDEX(" & sh_name & "_単!$I:$I,MATCH($F" & i & "," & sh_name & "_単!$A:$A,0)),"""")"
    Next i

    'H12～H18にindex matchの数式を入力  商品名
    For i = 12 To 18
        Range("H" & i).Formula = "=IFERROR(INDEX(" & sh_name & "_単!$C:$C,MATCH($F" & i & "," & sh_name & "_単!$A:$A,0)),"""")"
    Next i

    'J12～J18にindex matchの数式を入力  入数
    For i = 12 To 18
        Range("J" & i).Formula = "=IFERROR(INDEX(" & sh_name & "_単!$E:$E,MATCH($F" & i & "," & sh_name & "_単!$A:$A,0)),"""")"
    Next i

    'K12～K18にindex matchの数式を入力  数量
    For i = 12 To 18
        Range("K" & i).Formula = "=IFERROR(INDEX(" & sh_name & "_単!$F:$F,MATCH($F" & i & "," & sh_name & "_単!$A:$A,0)),"""")"
    Next i

    '---------------------------------------            混載            ---------------------------------------
    'G21～G27にindex matchの数式を入力  部門
    For i = 21 To 27
        Range("G" & i).Formula = "=IFERROR(INDEX(" & sh_name & "_混!$D:$D,MATCH($F" & i & "," & sh_name & "_混!$A:$A,0)),"""")"
    Next i

    'K21～K27にindex matchの数式を入力  数量
    For i = 21 To 27
        Range("K" & i).Formula = "=IFERROR(INDEX(" & sh_name & "_混!$E:$E,MATCH($F" & i & "," & sh_name & "_混!$A:$A,0)),"""")"
    Next i

    Application.ScreenUpdating = True                   '画面再開
    Application.Calculation = xlCalculationAutomatic    '計算再開
End Sub

Sub test3()
    Call 数式反映("大阪")
    Call 数式反映("小牧")

    Call 数式反映("仙台")
    Call 数式反映("青森")
    Call 数式反映("郡山")
End Sub
