Sub 非表示_列()
    For i = 2 To 22     ' B 列から V 列まで
        If Columns(i).Hidden = False Then
            Columns(i).Hidden = True ' 非表示
        End If
    Next i

    'BE列からCC列までを非表示にする
    For i = 31 To 49
        If Columns(i).Hidden = False Then
            Columns(i).Hidden = True ' 非表示
        End If
    Next i
    
    Range("X6").Select
End Sub

Sub 表示_列()
    For i = 2 To 22     ' B 列から V 列まで
        If Columns(i).Hidden = True Then
            Columns(i).Hidden = False ' 表示
        End If
    Next i

    'BE列からCC列までを表示にする
    For i = 31 To 49
        If Columns(i).Hidden = True Then
            Columns(i).Hidden = False ' 表示
        End If
    Next i
    
    Range("X6").Select
End Sub
