Attribute VB_Name = "�t�B���^�[_�`��2"
Sub �t�B���^�[�S���N���A_�`��2�V�[�g()
    Call �ی�.�S�ی����
    Worksheets("�`��2").Activate
    With ActiveSheet
        .Range("C5").Select
        If .FilterMode Then .ShowAllData
        
    End With
    Call �ی�.�����ی�
    MsgBox "�t�B���^�[�N���A����(�`��2�V�[�g)"
End Sub

Sub �t�B���^�[_�`��2()
    Call �ی�.�S�ی����
    Worksheets("�`��2").Activate
    
    With ActiveSheet
        .Range("C5").Select
        .Range("C5").AutoFilter Field:=25, Criteria1:="<>"
    End With
    Call �ی�.�����ی�
    MsgBox "�t�B���^�[�N���A����(�`��2�V�[�g)"
End Sub
