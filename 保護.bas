Sub �ی�()
    ActiveSheet.Protect Contents:=True, DrawingObjects:=False, UserInterfaceOnly:=True, AllowFormattingCells:=True
End Sub

Sub �ی����()
    ActiveSheet.Unprotect
End Sub

Sub �ی�_����()
    For ws_num = 1 To Worksheets.Count
        Worksheets(ws_num).Protect Contents:=True, DrawingObjects:=False, UserInterfaceOnly:=True, AllowFormattingCells:=True
    Next ws_num
    
End Sub

Sub �ی�_�S����()
    Dim sh As Object
    On Error Resume Next
    For Each sh In Sheets
    sh.Unprotect
    Next sh
    
End Sub