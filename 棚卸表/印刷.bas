Attribute VB_Name = "���"
'''''''''''''''''''''''''''''''          ���                     '''''''''''''''''''''''''''''''''''
Sub ���_�����ޗ�()
    Call �ی�.�S�ی����
    Worksheets("�����").Activate
    With ActiveSheet
        .Range("A10").Select
        If .FilterMode Then .ShowAllData
        .ListObjects("�e�[�u��2").Range.AutoFilter Field:=26, Criteria1:="<>"
        .ListObjects("�e�[�u��5").Range.AutoFilter Field:=1, Criteria1:="�����ޗ�"
    End With
    
    Call �\�[�g�����.�\�[�g�d���於
    
    '// �v�����^�Ƃ̐ڑ���ؒf
    Application.PrintCommunication = False
    '// ����ݒ�
    With ActiveSheet.PageSetup
        .PrintArea = Sheets("�����").Range("A8:J9016").Address
        .Orientation = xlPortrait '����������c�����ɐݒ�
        .Zoom = False '�g��k����ݒ�i���Ȃ��j
        .FitToPagesWide = 1 '���ׂĂ̗��1�y�[�W�Ɉ��
        .FitToPagesTall = False '�V�[�g��1�y�[�W�Ɉ��
        
        .LeftMargin = Application.CentimetersToPoints(0.5)
        .RightMargin = Application.CentimetersToPoints(0.5)
        .TopMargin = Application.CentimetersToPoints(0.5)
        .BottomMargin = Application.CentimetersToPoints(0.5)
        .HeaderMargin = Application.CentimetersToPoints(0)
        .FooterMargin = Application.CentimetersToPoints(0)
    End With
    '// �v�����^�Ɛڑ�����
    Application.PrintCommunication = True
    ActiveSheet.PrintOut
    
    Call ���_�ܖ�����
    Call �ی�.�����ی�
    Worksheets("�����").Activate
    ActiveSheet.Range("A11").Select
End Sub

Sub ���_���j��()
    Call �ی�.�S�ی����
    Worksheets("�����").Activate
    
    With ActiveSheet
        .Range("A10").Select
        If .FilterMode Then .ShowAllData
        .ListObjects("�e�[�u��2").Range.AutoFilter Field:=39, Criteria1:="<>"
        .ListObjects("�e�[�u��5").Range.AutoFilter Field:=1, Criteria1:="����"
    End With
    
    Call �\�[�g�����.�\�[�g�d���於
    
    '// �v�����^�Ƃ̐ڑ���ؒf
    Application.PrintCommunication = False
    '// ����ݒ�
    With ActiveSheet.PageSetup
        .PrintArea = Sheets("�����").Range("A8:J9019").Address
        .Orientation = xlPortrait '����������c�����ɐݒ�
        .Zoom = False '�g��k����ݒ�i���Ȃ��j
        .FitToPagesWide = 1 '���ׂĂ̗��1�y�[�W�Ɉ��
        .FitToPagesTall = False '�V�[�g��1�y�[�W�Ɉ��
        
        .LeftMargin = Application.CentimetersToPoints(0.5)
        .RightMargin = Application.CentimetersToPoints(0.5)
        .TopMargin = Application.CentimetersToPoints(0.5)
        .BottomMargin = Application.CentimetersToPoints(0.5)
        .HeaderMargin = Application.CentimetersToPoints(0)
        .FooterMargin = Application.CentimetersToPoints(0)
    End With
    '// �v�����^�Ɛڑ�����
    Application.PrintCommunication = True
    ActiveSheet.PrintOut
    Call �ی�.�����ی�
    ActiveSheet.Range("A11").Select
End Sub

Sub ���_���j��()
    Call �ی�.�S�ی����
    Worksheets("�����").Activate
    
    With ActiveSheet
        .Range("A10").Select
        If .FilterMode Then .ShowAllData
        .ListObjects("�e�[�u��2").Range.AutoFilter Field:=40, Criteria1:="<>"
        .ListObjects("�e�[�u��5").Range.AutoFilter Field:=1, Criteria1:="����"
    End With
    
    Call �\�[�g�����.�\�[�g�d���於
    
    '// �v�����^�Ƃ̐ڑ���ؒf
    Application.PrintCommunication = False
    '// ����ݒ�
    With ActiveSheet.PageSetup
        .PrintArea = Sheets("�����").Range("A8:J9019").Address
        .Orientation = xlPortrait '����������c�����ɐݒ�
        .Zoom = False '�g��k����ݒ�i���Ȃ��j
        .FitToPagesWide = 1 '���ׂĂ̗��1�y�[�W�Ɉ��
        .FitToPagesTall = False '�V�[�g��1�y�[�W�Ɉ��
        
        .LeftMargin = Application.CentimetersToPoints(0.5)
        .RightMargin = Application.CentimetersToPoints(0.5)
        .TopMargin = Application.CentimetersToPoints(0.5)
        .BottomMargin = Application.CentimetersToPoints(0.5)
        .HeaderMargin = Application.CentimetersToPoints(0)
        .FooterMargin = Application.CentimetersToPoints(0)
    End With
    '// �v�����^�Ɛڑ�����
    Application.PrintCommunication = True
    ActiveSheet.PrintOut
    Call �ی�.�����ی�
    ActiveSheet.Range("A11").Select
End Sub

Sub ���_�Ηj��()
    Call �ی�.�S�ی����
    Worksheets("�����").Activate
    
    With ActiveSheet
        .Range("A10").Select
        If .FilterMode Then .ShowAllData
        .ListObjects("�e�[�u��2").Range.AutoFilter Field:=41, Criteria1:="<>"
        .ListObjects("�e�[�u��5").Range.AutoFilter Field:=1, Criteria1:="����"
    End With
    
    Call �\�[�g�����.�\�[�g�d���於
    
    '// �v�����^�Ƃ̐ڑ���ؒf
    Application.PrintCommunication = False
    '// ����ݒ�
    With ActiveSheet.PageSetup
        .PrintArea = Sheets("�����").Range("A8:J9019").Address
        .Orientation = xlPortrait '����������c�����ɐݒ�
        .Zoom = False '�g��k����ݒ�i���Ȃ��j
        .FitToPagesWide = 1 '���ׂĂ̗��1�y�[�W�Ɉ��
        .FitToPagesTall = False '�V�[�g��1�y�[�W�Ɉ��
        
        .LeftMargin = Application.CentimetersToPoints(0.5)
        .RightMargin = Application.CentimetersToPoints(0.5)
        .TopMargin = Application.CentimetersToPoints(0.5)
        .BottomMargin = Application.CentimetersToPoints(0.5)
        .HeaderMargin = Application.CentimetersToPoints(0)
        .FooterMargin = Application.CentimetersToPoints(0)
    End With
    '// �v�����^�Ɛڑ�����
    Application.PrintCommunication = True
    ActiveSheet.PrintOut
    Call �ی�.�����ی�
    ActiveSheet.Range("A11").Select
End Sub

Sub ���_���j��()
    Call �ی�.�S�ی����
    Worksheets("�����").Activate
    
    With ActiveSheet
        .Range("A10").Select
        If .FilterMode Then .ShowAllData
        .ListObjects("�e�[�u��2").Range.AutoFilter Field:=42, Criteria1:="<>"
        .ListObjects("�e�[�u��5").Range.AutoFilter Field:=1, Criteria1:="����"
    End With
    
    Call �\�[�g�����.�\�[�g�d���於
    
    '// �v�����^�Ƃ̐ڑ���ؒf
    Application.PrintCommunication = False
    '// ����ݒ�
    With ActiveSheet.PageSetup
        .PrintArea = Sheets("�����").Range("A8:J9019").Address
        .Orientation = xlPortrait '����������c�����ɐݒ�
        .Zoom = False '�g��k����ݒ�i���Ȃ��j
        .FitToPagesWide = 1 '���ׂĂ̗��1�y�[�W�Ɉ��
        .FitToPagesTall = False '�V�[�g��1�y�[�W�Ɉ��
        
        .LeftMargin = Application.CentimetersToPoints(0.5)
        .RightMargin = Application.CentimetersToPoints(0.5)
        .TopMargin = Application.CentimetersToPoints(0.5)
        .BottomMargin = Application.CentimetersToPoints(0.5)
        .HeaderMargin = Application.CentimetersToPoints(0)
        .FooterMargin = Application.CentimetersToPoints(0)
    End With
    '// �v�����^�Ɛڑ�����
    Application.PrintCommunication = True
    ActiveSheet.PrintOut
    Call �ی�.�����ی�
    ActiveSheet.Range("A11").Select
End Sub

Sub ���_�ؗj��()
    Call �ی�.�S�ی����
    Worksheets("�����").Activate
    
    With ActiveSheet
        .Range("A10").Select
        If .FilterMode Then .ShowAllData
        .ListObjects("�e�[�u��2").Range.AutoFilter Field:=43, Criteria1:="<>"
        .ListObjects("�e�[�u��5").Range.AutoFilter Field:=1, Criteria1:="����"
    End With
    
    Call �\�[�g�����.�\�[�g�d���於
    
    '// �v�����^�Ƃ̐ڑ���ؒf
    Application.PrintCommunication = False
    '// ����ݒ�
    With ActiveSheet.PageSetup
        .PrintArea = Sheets("�����").Range("A8:J9019").Address
        .Orientation = xlPortrait '����������c�����ɐݒ�
        .Zoom = False '�g��k����ݒ�i���Ȃ��j
        .FitToPagesWide = 1 '���ׂĂ̗��1�y�[�W�Ɉ��
        .FitToPagesTall = False '�V�[�g��1�y�[�W�Ɉ��
        
        .LeftMargin = Application.CentimetersToPoints(0.5)
        .RightMargin = Application.CentimetersToPoints(0.5)
        .TopMargin = Application.CentimetersToPoints(0.5)
        .BottomMargin = Application.CentimetersToPoints(0.5)
        .HeaderMargin = Application.CentimetersToPoints(0)
        .FooterMargin = Application.CentimetersToPoints(0)
    End With
    '// �v�����^�Ɛڑ�����
    Application.PrintCommunication = True
    ActiveSheet.PrintOut
    Call �ی�.�����ی�
    ActiveSheet.Range("A11").Select
End Sub

Sub ���_���j��()
    Call �ی�.�S�ی����
    Worksheets("�����").Activate
    
    With ActiveSheet
        .Range("A10").Select
        If .FilterMode Then .ShowAllData
        .ListObjects("�e�[�u��2").Range.AutoFilter Field:=44, Criteria1:="<>"
        .ListObjects("�e�[�u��5").Range.AutoFilter Field:=1, Criteria1:="����"
    End With
    
    Call �\�[�g�����.�\�[�g�d���於
    
    '// �v�����^�Ƃ̐ڑ���ؒf
    Application.PrintCommunication = False
    '// ����ݒ�
    With ActiveSheet.PageSetup
        .PrintArea = Sheets("�����").Range("A8:J9019").Address
        .Orientation = xlPortrait '����������c�����ɐݒ�
        .Zoom = False '�g��k����ݒ�i���Ȃ��j
        .FitToPagesWide = 1 '���ׂĂ̗��1�y�[�W�Ɉ��
        .FitToPagesTall = False '�V�[�g��1�y�[�W�Ɉ��
        
        .LeftMargin = Application.CentimetersToPoints(0.5)
        .RightMargin = Application.CentimetersToPoints(0.5)
        .TopMargin = Application.CentimetersToPoints(0.5)
        .BottomMargin = Application.CentimetersToPoints(0.5)
        .HeaderMargin = Application.CentimetersToPoints(0)
        .FooterMargin = Application.CentimetersToPoints(0)
    End With
    '// �v�����^�Ɛڑ�����
    Application.PrintCommunication = True
    ActiveSheet.PrintOut
    Call �ی�.�����ی�
    ActiveSheet.Range("A11").Select
End Sub

Sub ���_�y�j��()
    Call �ی�.�S�ی����
    Worksheets("�����").Activate
    
    With ActiveSheet
        .Range("A10").Select
        If .FilterMode Then .ShowAllData
        .ListObjects("�e�[�u��2").Range.AutoFilter Field:=45, Criteria1:="<>"
        .ListObjects("�e�[�u��5").Range.AutoFilter Field:=1, Criteria1:="����"
    End With
    
    Call �\�[�g�����.�\�[�g�d���於
    
    '// �v�����^�Ƃ̐ڑ���ؒf
    Application.PrintCommunication = False
    '// ����ݒ�
    With ActiveSheet.PageSetup
        .PrintArea = Sheets("�����").Range("A8:J9019").Address
        .Orientation = xlPortrait '����������c�����ɐݒ�
        .Zoom = False '�g��k����ݒ�i���Ȃ��j
        .FitToPagesWide = 1 '���ׂĂ̗��1�y�[�W�Ɉ��
        .FitToPagesTall = False '�V�[�g��1�y�[�W�Ɉ��
        
        .LeftMargin = Application.CentimetersToPoints(0.5)
        .RightMargin = Application.CentimetersToPoints(0.5)
        .TopMargin = Application.CentimetersToPoints(0.5)
        .BottomMargin = Application.CentimetersToPoints(0.5)
        .HeaderMargin = Application.CentimetersToPoints(0)
        .FooterMargin = Application.CentimetersToPoints(0)
    End With
    '// �v�����^�Ɛڑ�����
    Application.PrintCommunication = True
    ActiveSheet.PrintOut
    Call �ی�.�����ی�
    ActiveSheet.Range("A11").Select
End Sub

Sub ���_IY()
    Call �ی�.�S�ی����
    Worksheets("�����").Activate
    
    With ActiveSheet
        .Range("A10").Select
        If .FilterMode Then .ShowAllData
        .ListObjects("�e�[�u��2").Range.AutoFilter Field:=31, Criteria1:="<>"
        .ListObjects("�e�[�u��5").Range.AutoFilter Field:=1, Criteria1:="IY"
    End With
    
    Call �\�[�g�����.�\�[�g�d���於
    
    '// �v�����^�Ƃ̐ڑ���ؒf
    Application.PrintCommunication = False
    '// ����ݒ�
    With ActiveSheet.PageSetup
        .PrintArea = Sheets("�����").Range("A8:J9023").Address
        .Orientation = xlPortrait '����������c�����ɐݒ�
        .Zoom = False '�g��k����ݒ�i���Ȃ��j
        .FitToPagesWide = 1 '���ׂĂ̗��1�y�[�W�Ɉ��
        .FitToPagesTall = False '�V�[�g��1�y�[�W�Ɉ��

        .LeftMargin = Application.CentimetersToPoints(0.5)
        .RightMargin = Application.CentimetersToPoints(0.5)
        .TopMargin = Application.CentimetersToPoints(0.5)
        .BottomMargin = Application.CentimetersToPoints(0.5)
        .HeaderMargin = Application.CentimetersToPoints(0)
        .FooterMargin = Application.CentimetersToPoints(0)
    End With
    '// �v�����^�Ɛڑ�����
    Application.PrintCommunication = True
    ActiveSheet.PrintOut
    Call �ی�.�����ی�
    ActiveSheet.Range("A11").Select
End Sub

Sub ���_CN�V�[�g()
    Call �ی�.�S�ی����
    Worksheets("���CN").Activate
    
    With ActiveSheet
        .Range("B4").Select
        If .FilterMode Then .ShowAllData
        .Range("B4").AutoFilter Field:=26, Criteria1:="<>"
    End With
    
    '// �v�����^�Ƃ̐ڑ���ؒf
    Application.PrintCommunication = False
    '// ����ݒ�
    With ActiveSheet.PageSetup
        .PrintArea = Sheets("���CN").Range("D4:V227").Address
        .Orientation = xlLandscape                                  '����������������ɐݒ�
        .Zoom = False                                               '�g��k����ݒ�i���Ȃ��j
        .FitToPagesWide = 1                                         '���ׂĂ̗��1�y�[�W�Ɉ��
        .FitToPagesTall = False                                     '�V�[�g��1�y�[�W�Ɉ��
        
        .LeftMargin = Application.CentimetersToPoints(0.5)
        .RightMargin = Application.CentimetersToPoints(0.5)
        .TopMargin = Application.CentimetersToPoints(0.5)
        .BottomMargin = Application.CentimetersToPoints(1)
        .HeaderMargin = Application.CentimetersToPoints(0)
        .FooterMargin = Application.CentimetersToPoints(0)
    End With
    '// �v�����^�Ɛڑ�����
    Application.PrintCommunication = True
    ActiveSheet.PrintOut
    Call �ی�.�����ی�
End Sub

Sub ���_�`��1()
    Call �ی�.�S�ی����
    Worksheets("�`��1").Activate
    
    Call �\�[�g�`��1.�\�[�g�`��1_ALL
    
    With ActiveSheet
        .ListObjects("�V��").Range.AutoFilter Field:=25, Criteria1:="<>"
        .ListObjects("���i�Ǘ�").Range.AutoFilter Field:=25, Criteria1:="<>"
        .ListObjects("�①��").Range.AutoFilter Field:=25, Criteria1:="<>"
        .ListObjects("�Ⓚ��").Range.AutoFilter Field:=25, Criteria1:="<>"
        .ListObjects("���̑�").Range.AutoFilter Field:=25, Criteria1:="<>"
    End With
    
    '// �v�����^�Ƃ̐ڑ���ؒf
    Application.PrintCommunication = False
    '// ����ݒ�
    With ActiveSheet.PageSetup
        .PrintArea = Sheets("�`��1").Range("D4:V216").Address
        .Orientation = xlLandscape                          '����������������ɐݒ�
        .Zoom = False                                       '�g��k����ݒ�i���Ȃ��j
        .FitToPagesWide = 1                                 '���ׂĂ̗��1�y�[�W�Ɉ��
        .FitToPagesTall = False                             '�V�[�g��1�y�[�W�Ɉ��

        .LeftMargin = Application.CentimetersToPoints(0.5)
        .RightMargin = Application.CentimetersToPoints(0.5)
        .TopMargin = Application.CentimetersToPoints(0.5)
        .BottomMargin = Application.CentimetersToPoints(1)
        .HeaderMargin = Application.CentimetersToPoints(0)
        .FooterMargin = Application.CentimetersToPoints(0)
    End With
    '// �v�����^�Ɛڑ�����
    Application.PrintCommunication = True
    ActiveSheet.PrintOut
    Call �ی�.�����ی�
    Worksheets("�����").Activate
    ActiveSheet.Range("A11").Select
End Sub

Sub ���_�`��2()
    Call �ی�.�S�ی����
    Worksheets("�`��2").Activate
    
    Call �\�[�g�`��2.�\�[�g�`��2_���V�s

    With ActiveSheet
        .Range("C5").Select
        .Range("C5").AutoFilter Field:=25, Criteria1:="<>"
    End With

    '// �v�����^�Ƃ̐ڑ���ؒf
    Application.PrintCommunication = False
    '// ����ݒ�
    With ActiveSheet.PageSetup
        .PrintArea = Sheets("�`��2").Range("D4:V73").Address
        .Orientation = xlLandscape                          '����������������ɐݒ�
        .Zoom = False                                       '�g��k����ݒ�i���Ȃ��j
        .FitToPagesWide = 1                                 '���ׂĂ̗��1�y�[�W�Ɉ��
        .FitToPagesTall = False                             '�V�[�g��1�y�[�W�Ɉ��

        .LeftMargin = Application.CentimetersToPoints(0.5)
        .RightMargin = Application.CentimetersToPoints(0.5)
        .TopMargin = Application.CentimetersToPoints(0.5)
        .BottomMargin = Application.CentimetersToPoints(1)
        .HeaderMargin = Application.CentimetersToPoints(0)
        .FooterMargin = Application.CentimetersToPoints(0)
    End With
    '// �v�����^�Ɛڑ�����
    Application.PrintCommunication = True
    ActiveSheet.PrintOut
    Call �ی�.�����ی�
    Worksheets("�����").Activate
    ActiveSheet.Range("A11").Select
End Sub

Sub ���_CN�܂Ƃ�()
    Call �ی�.�S�ی����
    Call ���_�`��1
    Call ���_�`��2
    Call �ی�.�����ی�
End Sub

Sub ���_�ܖ�����()
    Call �ی�.�S�ی����
    Worksheets("�ܖ�����").Activate
    
    With ActiveSheet
        .Range("A10").Select
        If .FilterMode Then .ShowAllData
        .Range("A10").AutoFilter Field:=1, Criteria1:="<>"
    End With
    
    '// �v�����^�Ƃ̐ڑ���ؒf
    Application.PrintCommunication = False
    '// ����ݒ�
    With ActiveSheet.PageSetup
        .PrintArea = Sheets("�ܖ�����").Range("A8:J55").Address
        .Orientation = xlPortrait '����������c�����ɐݒ�
        .Zoom = False '�g��k����ݒ�i���Ȃ��j
        .FitToPagesWide = 1 '���ׂĂ̗��1�y�[�W�Ɉ��
        .FitToPagesTall = False '�V�[�g��1�y�[�W�Ɉ��
        
        .LeftMargin = Application.CentimetersToPoints(0.5)
        .RightMargin = Application.CentimetersToPoints(0.5)
        .TopMargin = Application.CentimetersToPoints(0.5)
        .BottomMargin = Application.CentimetersToPoints(0.5)
        .HeaderMargin = Application.CentimetersToPoints(0)
        .FooterMargin = Application.CentimetersToPoints(0)
    End With
    '// �v�����^�Ɛڑ�����
    Application.PrintCommunication = True
    ActiveSheet.PrintOut
    Call �ی�.�����ی�
End Sub

Sub ���_�V�[��()
    Call �ی�.�S�ی����
    Worksheets("�����").Activate
    
    With ActiveSheet
        .Range("A10").Select
        If .FilterMode Then .ShowAllData
        .ListObjects("�e�[�u��2").Range.AutoFilter Field:=38, Criteria1:="�V�[��"
        .ListObjects("�e�[�u��2").Range.AutoFilter Field:=6, Criteria1:="<>-"
        .ListObjects("�e�[�u��5").Range.AutoFilter Field:=1, Criteria1:="�V�[��"
    End With
    
    Call �\�[�g�����.�\�[�g�d���於
    
    '// �v�����^�Ƃ̐ڑ���ؒf
    Application.PrintCommunication = False
    '// ����ݒ�
    With ActiveSheet.PageSetup
        .PrintArea = Sheets("�����").Range("A8:J9026").Address
        .Orientation = xlPortrait '����������c�����ɐݒ�
        .Zoom = False '�g��k����ݒ�i���Ȃ��j
        .FitToPagesWide = 1 '���ׂĂ̗��1�y�[�W�Ɉ��
        .FitToPagesTall = False '�V�[�g��1�y�[�W�Ɉ��
        
        .LeftMargin = Application.CentimetersToPoints(0.5)
        .RightMargin = Application.CentimetersToPoints(0.5)
        .TopMargin = Application.CentimetersToPoints(0.5)
        .BottomMargin = Application.CentimetersToPoints(0.5)
        .HeaderMargin = Application.CentimetersToPoints(0)
        .FooterMargin = Application.CentimetersToPoints(0)
    End With
    '// �v�����^�Ɛڑ�����
    Application.PrintCommunication = True
    ActiveSheet.PrintOut
    Call �ی�.�����ی�
    ActiveSheet.Range("A11").Select
End Sub

Sub ���_�g���[()
    Call �ی�.�S�ی����
    Worksheets("�g���[").Activate

    '// �v�����^�Ƃ̐ڑ���ؒf
    Application.PrintCommunication = False
    '// ����ݒ�
    With ActiveSheet.PageSetup
        '.PrintArea = Sheets("���CN").Range("D4:V183").Address
        .Orientation = xlPortrait                                  '����������c�����ɐݒ�
        .Zoom = False                                               '�g��k����ݒ�i���Ȃ��j
        .FitToPagesWide = 1                                         '���ׂĂ̗��1�y�[�W�Ɉ��
        .FitToPagesTall = False                                     '�V�[�g��1�y�[�W�Ɉ��
        
        .LeftMargin = Application.CentimetersToPoints(1)
        .RightMargin = Application.CentimetersToPoints(1)
        .TopMargin = Application.CentimetersToPoints(1)
        .BottomMargin = Application.CentimetersToPoints(1)
        .HeaderMargin = Application.CentimetersToPoints(1)
        .FooterMargin = Application.CentimetersToPoints(0)
    End With
    '// �v�����^�Ɛڑ�����
    Application.PrintCommunication = True
    ActiveSheet.PrintOut
    Call �ی�.�����ی�
End Sub

Sub ���_������()
    '�����I����p�̃{�^�� 2020.09.13�R�����g
    '�ڍׂ́u�t�B���^�[_������V�[�g�v�́u�t�B���^�[�������v�֐��ɃR�����g�ŏ����Ă���

    Call �ی�.�S�ی����
    Worksheets("�����").Activate
    With ActiveSheet
        .Range("A10").Select
        If .FilterMode Then .ShowAllData
        .ListObjects("�e�[�u��2").Range.AutoFilter Field:=11, Criteria1:="<>"
        .ListObjects("�e�[�u��2").Range.AutoFilter Field:=20, Criteria1:=""
        .ListObjects("�e�[�u��2").Range.AutoFilter Field:=27, Criteria1:=""
        .ListObjects("�e�[�u��2").Range.AutoFilter Field:=28, Criteria1:=""
        .ListObjects("�e�[�u��2").Range.AutoFilter Field:=62, Criteria1:="��"
        .ListObjects("�e�[�u��5").Range.AutoFilter Field:=1, Criteria1:="����"
    End With
    
    Call �\�[�g�����.�\�[�g�d���於
    
    '// �v�����^�Ƃ̐ڑ���ؒf
    Application.PrintCommunication = False
    '// ����ݒ�
    With ActiveSheet.PageSetup
        .PrintArea = Sheets("�����").Range("A8:J9016").Address
        .Orientation = xlPortrait '����������c�����ɐݒ�
        .Zoom = False '�g��k����ݒ�i���Ȃ��j
        .FitToPagesWide = 1 '���ׂĂ̗��1�y�[�W�Ɉ��
        .FitToPagesTall = False '�V�[�g��1�y�[�W�Ɉ��
        
        .LeftMargin = Application.CentimetersToPoints(0.5)
        .RightMargin = Application.CentimetersToPoints(0.5)
        .TopMargin = Application.CentimetersToPoints(0.5)
        .BottomMargin = Application.CentimetersToPoints(0.5)
        .HeaderMargin = Application.CentimetersToPoints(0)
        .FooterMargin = Application.CentimetersToPoints(0)
    End With
    '// �v�����^�Ɛڑ�����
    Application.PrintCommunication = True
    ActiveSheet.PrintOut
    
    'Call ���_�ܖ�����
    Call �ی�.�����ی�
    Worksheets("�����").Activate
    ActiveSheet.Range("A11").Select
End Sub
