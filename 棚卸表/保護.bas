Attribute VB_Name = "�ی�"
Sub �ی�()
    ActiveSheet.Protect AllowFiltering:=True
End Sub

Sub �ی����()
    ActiveSheet.Unprotect
End Sub

Sub �����ی�()
    Worksheets("�ݒ�").Protect AllowFiltering:=True
    Worksheets("�����").Protect AllowFiltering:=True
    Worksheets("�ܖ�����").Protect AllowFiltering:=True
    Worksheets("���CN").Protect AllowFiltering:=True
    Worksheets("�`��1").Protect AllowFiltering:=True
    Worksheets("�`��2").Protect AllowFiltering:=True
End Sub

Sub �S�ی����()
    Dim sh As Object
    On Error Resume Next
    For Each sh In Sheets
    sh.Unprotect
    Next sh
End Sub
