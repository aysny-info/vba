Attribute VB_Name = "�\�[�g�����"
Sub �\�[�g�d���於()
    Call �ی�.�S�ی����
    Worksheets("�����").Activate
    
    ActiveWorkbook.Worksheets("�����").ListObjects("�e�[�u��2").Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("�����").ListObjects("�e�[�u��2").Sort.SortFields.Add Key _
        :=Range("�e�[�u��2[[#All],[�d���於]]"), SortOn:=xlSortOnValues, Order:= _
        xlAscending, DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets("�����").ListObjects("�e�[�u��2").Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    Call �ی�.�����ی�
End Sub

Sub �\�[�g���i�R�[�h()
    Call �ی�.�S�ی����
    Worksheets("�����").Activate
    
    ActiveWorkbook.Worksheets("�����").ListObjects("�e�[�u��2").Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("�����").ListObjects("�e�[�u��2").Sort.SortFields.Add Key _
        :=Range("�e�[�u��2[[#All],[���i�R�[�h]]"), SortOn:=xlSortOnValues, Order:= _
        xlAscending, DataOption:=xlSortNormal       'xlAscending ����     xlDescending �~��
    With ActiveWorkbook.Worksheets("�����").ListObjects("�e�[�u��2").Sort
        .Header = xlYes                          '�擪�s�����o���Ƃ��Ďg�p
        .MatchCase = False                      '�啶������������ʂ��Ȃ�
        .Orientation = xlTopToBottom            '�s�P�ʂŕ��בւ�
        .SortMethod = xlPinYin                  '�ӂ肪�Ȃ��g��Ȃ�
        .Apply                                  '���בւ������s
    End With
    Call �ی�.�����ی�
End Sub

Sub �\�[�g���W()
    Call �ی�.�S�ی����
    Worksheets("�����").Activate
    
    ActiveWorkbook.Worksheets("�����").ListObjects("�e�[�u��2").Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("�����").ListObjects("�e�[�u��2").Sort.SortFields.Add Key _
        :=Range("�e�[�u��2[[#All],[���W]]"), SortOn:=xlSortOnValues, Order:= _
        xlAscending, DataOption:=xlSortNormal       'xlAscending ����     xlDescending �~��
    With ActiveWorkbook.Worksheets("�����").ListObjects("�e�[�u��2").Sort
        .Header = xlYes                          '�擪�s�����o���Ƃ��Ďg�p
        .MatchCase = False                      '�啶������������ʂ��Ȃ�
        .Orientation = xlTopToBottom            '�s�P�ʂŕ��בւ�
        .SortMethod = xlPinYin                  '�ӂ肪�Ȃ��g��Ȃ�
        .Apply                                  '���בւ������s
    End With
    Call �ی�.�����ی�
End Sub

Sub �\�[�g�N���A()
    Call �ی�.�S�ی����
    ActiveWorkbook.Worksheets("�����").ListObjects("�e�[�u��2").Sort.SortFields.Clear
    Call �ی�.�����ی�
End Sub

Sub �\�[�g���Z�b�g()
    Call �\�[�g���i�R�[�h
    Call �\�[�g�N���A
    Worksheets("�����").Activate
    MsgBox "�\�[�g���Z�b�g����"
End Sub
