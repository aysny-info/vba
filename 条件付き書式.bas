Sub 条件付き書式_追加()
    ThisWorkbook.Activate
    Worksheets("Sheet1").Activate
    
    'Range("$B:$D").FormatConditions.Delete      '条件付き書式の削除。追加する前に削除したほうがいい場合は削除。他の条件付き書式と被る場合はコメントアウト。
    Dim fc As FormatCondition
    Set fc = Range("$B:$D").FormatConditions.Add(Type:=xlExpression, Formula1:="=$A1=""あいうえお""")   'A1に「あいうえお」が入っている場合
    'Set fc = Range("$B:$D").FormatConditions.Add(Type:=xlExpression, Formula1:="=left($A1,1)=""Ｃ""")   'A1の左１文字が「Ｃ」の場合
    'Set fc = Range("$B:$D").FormatConditions.Add(Type:=xlExpression, Formula1:="=$A1>=1")             'A1が1以上の場合

    fc.Font.Color = RGB(255, 255, 255) '白 フォントカラー
    fc.Interior.Color = RGB(0, 0, 0)   '黒 背景色
End Sub

Sub 条件付き書式_削除()
    Range("$B:$D").FormatConditions.Delete      '条件付き書式の削除
End Sub