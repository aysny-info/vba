
Sub ���T�̌v��ǂݍ���()
    '�A�N�e�B�u
    Worksheets("�s�b�L���O�\").Activate
    Worksheets("�s�b�L���O�\").Select
    
    Dim rc As VbMsgBoxResult
    rc = MsgBox("���O���͂̏��������܂��B" & vbCrLf & "�̔��v��W�v�\�Łu���T�v�ɐݒ肵�����i��C��ɓ\��t���܂����H", vbYesNo + vbQuestion)
    If rc = vbNo Then
        MsgBox "���O���͂̏����𒆎~���܂�", vbCritical
        Exit Sub
    End If
    Worksheets("�s�b�L���O�\").Unprotect   '�ی����
    
    modosu_ship = Worksheets("�s�b�L���O�\").Range("D6")   '�߂��p
    Application.ScreenUpdating = False                 '��ʒ�~
    Application.Calculation = xlCalculationManual     '�蓮�v�Z
    
    CalenderForm.Show   '�J�����_�[
    'Application.Calculate   '�Čv�Z
    kakunin_ship = Worksheets("�s�b�L���O�\").Range("D6")  '�o�ד��擾
    
    rc = MsgBox(Str(kakunin_ship) + "�̏��i�ꗗ��ǂݍ��݁A���ʂ��N���A���܂����H", vbYesNo + vbQuestion)
    If rc = vbNo Then
        Worksheets("�s�b�L���O�\").Range("D6") = modosu_ship    '�߂�
    Application.Calculation = xlCalculationAutomatic    '�����v�Z
    Application.ScreenUpdating = True                 '���
    Worksheets("�s�b�L���O�\").Protect   '�ی�
        
        MsgBox "���O���͂̏������𒆎~���܂�", vbCritical
        Exit Sub
    End If
    
    '�A�N�e�B�u
    Worksheets("�v��").Activate
    Worksheets("�v��").Select
    
    'AU����擾���ď��i�̐����擾
    au_column = Range(Cells(3, 47), Cells(100, 47))
    max_I = 1
    For i = 1 To UBound(au_column)
        'If Left(AU_column(i, 1), 1) <> "�h" Then
        If IsError(au_column(i, 1)) Then
            max_I = i - 1
            Exit For
        End If
    Next i
    
    '�A�N�e�B�u
    Worksheets("�s�b�L���O�\").Activate
    Worksheets("�s�b�L���O�\").Select
    
    
    '���ʍ폜��C��폜
    Range("F8:I38,K8:L38").ClearContents
    Range(Cells(8, 3), Cells(38, 3)).ClearContents
    'C��\�t��
    Range(Cells(8, 3), Cells(8 + max_I - 1, 3)) = au_column
    
    
    Application.Calculation = xlCalculationAutomatic    '�����v�Z
    Application.ScreenUpdating = True                 '���
    Worksheets("�s�b�L���O�\").Protect   '�ی�
    
    MsgBox ("���O���͂̏������������܂����B" & vbCrLf & "���ʓ��͌�A�A���O���͂̕ۑ�(�o��)�{�^���������Ă��������B")
End Sub

