Attribute VB_Name = "�t�B���^�[_�`��1"
Sub �t�B���^�[�S���N���A_�`��1�V�[�g()
    Call �ی�.�S�ی����
    Worksheets("�`��1").Activate
    With ActiveSheet
        .Range("C5").Select
        If .FilterMode Then .ShowAllData
        
        .Range("C37").Select
        If .FilterMode Then .ShowAllData
        
        .Range("C69").Select
        If .FilterMode Then .ShowAllData
        
        .Range("C101").Select
        If .FilterMode Then .ShowAllData
        
        .Range("C163").Select
        If .FilterMode Then .ShowAllData
        
        .Range("C5").Select
    End With
    Call �ی�.�����ی�
    MsgBox "�t�B���^�[�N���A����(�`��1�V�[�g)"
End Sub

Sub �t�B���^�[_�`��1()
    Call �ی�.�S�ی����
    Worksheets("�`��1").Activate
    
    With ActiveSheet
        .ListObjects("�V��").Range.AutoFilter Field:=25, Criteria1:="<>"
        .ListObjects("���i�Ǘ�").Range.AutoFilter Field:=25, Criteria1:="<>"
        .ListObjects("�①��").Range.AutoFilter Field:=25, Criteria1:="<>"
        .ListObjects("�Ⓚ��").Range.AutoFilter Field:=25, Criteria1:="<>"
        .ListObjects("���̑�").Range.AutoFilter Field:=25, Criteria1:="<>"
    End With
    Call �ی�.�����ی�
    MsgBox "�t�B���^�[�N���A����(�`��1�V�[�g)"
End Sub
    
