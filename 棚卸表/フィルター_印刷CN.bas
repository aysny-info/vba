Attribute VB_Name = "�t�B���^�[_���CN"
'''''''''''''''''''''''''''''''      �t�B���^�[�֘A                   ''''''''''''''''''''''''''''

Sub �t�B���^�[�S���N���A_���CN�V�[�g()
    Call �ی�.�S�ی����
    Worksheets("���CN").Activate
    With ActiveSheet
        .Range("B4").Select
        If .FilterMode Then .ShowAllData
    End With
    Call �ی�.�����ی�
    MsgBox "�t�B���^�[�N���A����(CN����)"
End Sub

Sub �t�B���^�[CN_���CN�V�[�g()
    Call �ی�.�S�ی����
    Worksheets("���CN").Activate
    With ActiveSheet
        .Range("B4").Select
        If .FilterMode Then .ShowAllData
        .Range("B4").AutoFilter Field:=26, Criteria1:="<>"
    End With
    Call �ی�.�����ی�
    MsgBox "�t�B���^�[����(CN����)"
End Sub

