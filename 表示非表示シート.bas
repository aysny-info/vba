Sub 表示_シート全()

   Dim WS As Worksheet
    For Each WS In Worksheets
        WS.Visible = True
    Next
    
End Sub

Sub 非表示_シート()
   Dim WS As Worksheet, flag As Boolean
    For Each WS In Worksheets
        If WS.Name = "" Or _
           WS.Name = "" Or _
           WS.Name = "" _
        Then
        Else: WS.Visible = False
        End If
    Next
   
End Sub