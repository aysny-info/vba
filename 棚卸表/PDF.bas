Attribute VB_Name = "PDF"
'''''''''''''''''''''''''''''''          PDF                     '''''''''''''''''''''''''''''''''''
Sub PDF�����ޗ�()
    Call �ی�.�S�ی����
    
    Worksheets("�����").Activate
    
    With ActiveSheet
        .Range("A10").Select
        If .FilterMode Then .ShowAllData
        .ListObjects("�e�[�u��2").Range.AutoFilter Field:=26, Criteria1:="<>"
        .ListObjects("�e�[�u��5").Range.AutoFilter Field:=1, Criteria1:="�����ޗ�"
    End With
    
    Call �\�[�g�����.�\�[�g�d���於
    
    'PDF�ݒ�
    Dim fileName As String '�ۑ���t�H���_�p�X���t�@�C����
    hiduke = Format(Month(Now), "00") & Format(Day(Now), "00") & "_" & Format(Hour(Now), "00") & Format(Minute(Now), "00") & Format(Second(Now), "00")
    
    fileName = ThisWorkbook.Path & "\PDF\" & "�����ޗ�" & hiduke & ".pdf"

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
    
    With ActiveSheet
        .ExportAsFixedFormat Type:=xlTypePDF, fileName:=fileName
    End With
    Call PDF_�ܖ�����
    Call �ی�.�����ی�
    
    MsgBox "�����ޗ��̈������"
End Sub
Sub PDF���j��()
    Call �ی�.�S�ی����
    Worksheets("�����").Activate
    
    With ActiveSheet
        .Range("A10").Select
        If .FilterMode Then .ShowAllData
        .ListObjects("�e�[�u��2").Range.AutoFilter Field:=39, Criteria1:="<>"
        .ListObjects("�e�[�u��5").Range.AutoFilter Field:=1, Criteria1:="����"
    End With
    
    Call �\�[�g�����.�\�[�g�d���於
    
    'PDF�ݒ�
    Dim fileName As String '�ۑ���t�H���_�p�X���t�@�C����
    hiduke = Format(Month(Now), "00") & Format(Day(Now), "00") & "_" & Format(Hour(Now), "00") & Format(Minute(Now), "00") & Format(Second(Now), "00")
    
    fileName = ThisWorkbook.Path & "\PDF\" & "����(��)" & hiduke & ".pdf"

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
    
    With ActiveSheet
        .ExportAsFixedFormat Type:=xlTypePDF, fileName:=fileName
    End With
    Call �ی�.�����ی�
    MsgBox "���j����PDF����"
End Sub

Sub PDF���j��()
    Call �ی�.�S�ی����
    Worksheets("�����").Activate
    
    With ActiveSheet
        .Range("A10").Select
        If .FilterMode Then .ShowAllData
        .ListObjects("�e�[�u��2").Range.AutoFilter Field:=40, Criteria1:="<>"
        .ListObjects("�e�[�u��5").Range.AutoFilter Field:=1, Criteria1:="����"
    End With
    
    Call �\�[�g�����.�\�[�g�d���於
    
    'PDF�ݒ�
    Dim fileName As String '�ۑ���t�H���_�p�X���t�@�C����
    hiduke = Format(Month(Now), "00") & Format(Day(Now), "00") & "_" & Format(Hour(Now), "00") & Format(Minute(Now), "00") & Format(Second(Now), "00")
    
    fileName = ThisWorkbook.Path & "\PDF\" & "����(��)" & hiduke & ".pdf"
    
    
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
    
    With ActiveSheet
        .ExportAsFixedFormat Type:=xlTypePDF, fileName:=fileName
    End With
    Call �ی�.�����ی�
    MsgBox "���j����PDF����"
End Sub

Sub PDF�Ηj��()
    Call �ی�.�S�ی����
    Worksheets("�����").Activate
    
    With ActiveSheet
        .Range("A10").Select
        If .FilterMode Then .ShowAllData
        .ListObjects("�e�[�u��2").Range.AutoFilter Field:=41, Criteria1:="<>"
        .ListObjects("�e�[�u��5").Range.AutoFilter Field:=1, Criteria1:="����"
    End With
    
    Call �\�[�g�����.�\�[�g�d���於
    
    'PDF�ݒ�
    Dim fileName As String '�ۑ���t�H���_�p�X���t�@�C����
    hiduke = Format(Month(Now), "00") & Format(Day(Now), "00") & "_" & Format(Hour(Now), "00") & Format(Minute(Now), "00") & Format(Second(Now), "00")
    
    fileName = ThisWorkbook.Path & "\PDF\" & "����(��)" & hiduke & ".pdf"
    
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
    
    With ActiveSheet
        .ExportAsFixedFormat Type:=xlTypePDF, fileName:=fileName
    End With
    Call �ی�.�����ی�
    MsgBox "�Ηj����PDF����"
End Sub

Sub PDF���j��()
    Call �ی�.�S�ی����
    Worksheets("�����").Activate
    
    With ActiveSheet
        .Range("A10").Select
        If .FilterMode Then .ShowAllData
        .ListObjects("�e�[�u��2").Range.AutoFilter Field:=42, Criteria1:="<>"
        .ListObjects("�e�[�u��5").Range.AutoFilter Field:=1, Criteria1:="����"
    End With
    
    Call �\�[�g�����.�\�[�g�d���於
    
    'PDF�ݒ�
    Dim fileName As String '�ۑ���t�H���_�p�X���t�@�C����
    hiduke = Format(Month(Now), "00") & Format(Day(Now), "00") & "_" & Format(Hour(Now), "00") & Format(Minute(Now), "00") & Format(Second(Now), "00")
    
    fileName = ThisWorkbook.Path & "\PDF\" & "����(��)" & hiduke & ".pdf"
    
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
    
    With ActiveSheet
        .ExportAsFixedFormat Type:=xlTypePDF, fileName:=fileName
    End With
    Call �ی�.�����ی�
    MsgBox "���j����PDF����"
End Sub

Sub PDF�ؗj��()
    Call �ی�.�S�ی����
    Worksheets("�����").Activate
    
    With ActiveSheet
        .Range("A10").Select
        If .FilterMode Then .ShowAllData
        .ListObjects("�e�[�u��2").Range.AutoFilter Field:=43, Criteria1:="<>"
        .ListObjects("�e�[�u��5").Range.AutoFilter Field:=1, Criteria1:="����"
    End With
    
    Call �\�[�g�����.�\�[�g�d���於
    
    'PDF�ݒ�
    Dim fileName As String '�ۑ���t�H���_�p�X���t�@�C����
    hiduke = Format(Month(Now), "00") & Format(Day(Now), "00") & "_" & Format(Hour(Now), "00") & Format(Minute(Now), "00") & Format(Second(Now), "00")
    
    fileName = ThisWorkbook.Path & "\PDF\" & "����(��)" & hiduke & ".pdf"
    
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
    
    With ActiveSheet
        .ExportAsFixedFormat Type:=xlTypePDF, fileName:=fileName
    End With
    Call �ی�.�����ی�
    MsgBox "�ؗj����PDF����"
End Sub

Sub PDF���j��()
    Call �ی�.�S�ی����
    Worksheets("�����").Activate
    
    With ActiveSheet
        .Range("A10").Select
        If .FilterMode Then .ShowAllData
        .ListObjects("�e�[�u��2").Range.AutoFilter Field:=44, Criteria1:="<>"
        .ListObjects("�e�[�u��5").Range.AutoFilter Field:=1, Criteria1:="����"
    End With
    
    Call �\�[�g�����.�\�[�g�d���於
    
    'PDF�ݒ�
    Dim fileName As String '�ۑ���t�H���_�p�X���t�@�C����
    hiduke = Format(Month(Now), "00") & Format(Day(Now), "00") & "_" & Format(Hour(Now), "00") & Format(Minute(Now), "00") & Format(Second(Now), "00")
    
    fileName = ThisWorkbook.Path & "\PDF\" & "����(��)" & hiduke & ".pdf"
    
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
    
    With ActiveSheet
        .ExportAsFixedFormat Type:=xlTypePDF, fileName:=fileName
    End With
    Call �ی�.�����ی�
    MsgBox "���j����PDF����"
End Sub

Sub PDF�y�j��()
    Call �ی�.�S�ی����
    Worksheets("�����").Activate
    
    With ActiveSheet
        .Range("A10").Select
        If .FilterMode Then .ShowAllData
        .ListObjects("�e�[�u��2").Range.AutoFilter Field:=45, Criteria1:="<>"
        .ListObjects("�e�[�u��5").Range.AutoFilter Field:=1, Criteria1:="����"
    End With
    
    Call �\�[�g�����.�\�[�g�d���於
    
    'PDF�ݒ�
    Dim fileName As String '�ۑ���t�H���_�p�X���t�@�C����
    hiduke = Format(Month(Now), "00") & Format(Day(Now), "00") & "_" & Format(Hour(Now), "00") & Format(Minute(Now), "00") & Format(Second(Now), "00")
    
    fileName = ThisWorkbook.Path & "\PDF\" & "����(�y)" & hiduke & ".pdf"
    
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
    
    With ActiveSheet
        .ExportAsFixedFormat Type:=xlTypePDF, fileName:=fileName
    End With
    Call �ی�.�����ی�
    MsgBox "�y�j����PDF����"
End Sub

Sub PDF_IY()
    Call �ی�.�S�ی����
    Worksheets("�����").Activate
    With ActiveSheet
        .Range("A10").Select
        If .FilterMode Then .ShowAllData
        .ListObjects("�e�[�u��2").Range.AutoFilter Field:=31, Criteria1:="<>"
        .ListObjects("�e�[�u��5").Range.AutoFilter Field:=1, Criteria1:="IY"
    End With
    
    Call �\�[�g�����.�\�[�g�d���於
    
    'PDF�ݒ�
    Dim fileName As String '�ۑ���t�H���_�p�X���t�@�C����
    hiduke = Format(Month(Now), "00") & Format(Day(Now), "00") & "_" & Format(Hour(Now), "00") & Format(Minute(Now), "00") & Format(Second(Now), "00")
    
    fileName = ThisWorkbook.Path & "\PDF\" & "IY_" & hiduke & ".pdf"
    
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
    
    With ActiveSheet
        .ExportAsFixedFormat Type:=xlTypePDF, fileName:=fileName
    End With
    Call �ی�.�����ی�
    MsgBox "IY��PDF����"
End Sub

Sub PDF_CN()
    Call �ی�.�S�ی����
    ''''''''''''''''''''''''''''''''''''''''''''''   �`���P  '''''''''''''''''''''''''''''''''''''''''''''
    Worksheets("�`��1").Activate
    
    Call �\�[�g�`��1.�\�[�g�`��1_ALL
    
    With ActiveSheet
        .ListObjects("�V��").Range.AutoFilter Field:=25, Criteria1:="<>"
        .ListObjects("���i�Ǘ�").Range.AutoFilter Field:=25, Criteria1:="<>"
        .ListObjects("�①��").Range.AutoFilter Field:=25, Criteria1:="<>"
        .ListObjects("�Ⓚ��").Range.AutoFilter Field:=25, Criteria1:="<>"
        .ListObjects("���̑�").Range.AutoFilter Field:=25, Criteria1:="<>"
    End With
        
    'PDF�ݒ�
    Dim fileName As String '�ۑ���t�H���_�p�X���t�@�C����
    hiduke = Format(Month(Now), "00") & Format(Day(Now), "00") & "_" & Format(Hour(Now), "00") & Format(Minute(Now), "00") & Format(Second(Now), "00")
    
    fileName = ThisWorkbook.Path & "\PDF\" & "�R�[�vNo1_" & hiduke & ".pdf"
    
    With ActiveSheet.PageSetup
        .PrintArea = Sheets("�`��1").Range("D4:V183").Address
        .Orientation = xlLandscape '����������������ɐݒ�
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
    
    With ActiveSheet
        .ExportAsFixedFormat Type:=xlTypePDF, fileName:=fileName
    End With
    
    ''''''''''''''''''''''''''''''''''''''''''''''   �`���Q  '''''''''''''''''''''''''''''''''''''''''''''
    Worksheets("�`��2").Activate
    
    Call �\�[�g�`��2.�\�[�g�`��2_���V�s

    With ActiveSheet
        .Range("C5").Select
        .Range("C5").AutoFilter Field:=25, Criteria1:="<>"
    End With
    
    'PDF�ݒ�
    Dim fileName2 As String '�ۑ���t�H���_�p�X���t�@�C����
    
    fileName2 = ThisWorkbook.Path & "\PDF\" & "�R�[�vNo2(���V�s)_" & hiduke & ".pdf"
    
    With ActiveSheet.PageSetup
        .PrintArea = Sheets("�`��2").Range("D4:V58").Address
        .Orientation = xlLandscape '����������������ɐݒ�
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
    
    With ActiveSheet
        .ExportAsFixedFormat Type:=xlTypePDF, fileName:=fileName2
    End With
    Call �ی�.�����ی�
    Worksheets("�����").Activate
    MsgBox "CN��PDF����"
End Sub

Sub PDF_�ܖ�����()
    Call �ی�.�S�ی����
    Worksheets("�ܖ�����").Activate
    
    With ActiveSheet
        .Range("A10").Select
        If .FilterMode Then .ShowAllData
        .Range("A10").AutoFilter Field:=1, Criteria1:="<>"
    End With
    
    'PDF�ݒ�
    Dim fileName As String '�ۑ���t�H���_�p�X���t�@�C����
    hiduke = Format(Month(Now), "00") & Format(Day(Now), "00") & "_" & Format(Hour(Now), "00") & Format(Minute(Now), "00") & Format(Second(Now), "00")
    
    fileName = ThisWorkbook.Path & "\PDF\" & "�ܖ�����" & hiduke & ".pdf"
    
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
    
    With ActiveSheet
        .ExportAsFixedFormat Type:=xlTypePDF, fileName:=fileName
    End With
    Call �ی�.�����ی�
    MsgBox "�ܖ�������PDF����"
End Sub

Sub PDF�V�[��()
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
    
    'PDF�ݒ�
    Dim fileName As String '�ۑ���t�H���_�p�X���t�@�C����
    hiduke = Format(Month(Now), "00") & Format(Day(Now), "00") & "_" & Format(Hour(Now), "00") & Format(Minute(Now), "00") & Format(Second(Now), "00")
    
    fileName = ThisWorkbook.Path & "\PDF\" & "�V�[��" & hiduke & ".pdf"
    
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
    
    With ActiveSheet
        .ExportAsFixedFormat Type:=xlTypePDF, fileName:=fileName
    End With
    Call �ی�.�����ی�
    MsgBox "�V�[����PDF����"
End Sub
