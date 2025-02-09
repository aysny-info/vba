Sub 原料展開_数式を値に()
    Dim wsDst As Worksheet
    Dim srcArray As Variant
    Dim 最終行 As Long
    
    Set wsDst = ThisWorkbook.Sheets("test")
    
    最終行 = 10
    
    ' オートフィルで最終行までコピー ---
    wsDst.Range("A3:BD3").AutoFill Destination:=wsDst.Range("A3:BD" & 最終行)
    
    ' --- 値貼り付けで数式を除去 ---
    With wsDst.Range("A4:BD" & 最終行)
        .Value = .Value
    End With
    
    MsgBox "完了しました。"

End Sub


