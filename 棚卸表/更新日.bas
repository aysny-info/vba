Attribute VB_Name = "�X�V��"
Sub �ܖ�����_COPY_OPEN()
    Call �ی�.�S�ی����

    Worksheets("�����").Activate
    
    With ActiveSheet
        .Range("A10").Select
        If .FilterMode Then .ShowAllData
    End With
    
    Dim y As Long, i As Long, j As Long, v As Long, syomi As Variant
    syomi = Range(Cells(11, 13), Cells(9010, 18))           ''�ܖ������X�V�O
    Range(Cells(11, 46), Cells(9010, 51)) = syomi           ''�y�[�X�g
End Sub

Sub �X�V��_����()
    Call �ی�.�S�ی����

    Worksheets("�����").Activate
    
    With ActiveSheet
        .Range("A10").Select
        If .FilterMode Then .ShowAllData
    End With

    Dim koushin As Variant, syo_mi_mae As Variant, syo_mi_ato As Variant
    Dim i As Long, s As Long, d As Date
    d = Date
    syo_mi_mae = Range(Cells(11, 46), Cells(9010, 51))     ''�ܖ������ύX�O
    syo_mi_ato = Range(Cells(11, 13), Cells(9010, 18))     ''�ܖ������ύX��
    koushin = Range(Cells(11, 12), Cells(9010, 12))        ''�X�V��
    
    For i = 1 To 8999                                      ''��r����
        For s = 1 To 6
            If syo_mi_mae(i, s) <> syo_mi_ato(i, s) Then
                koushin(i, 1) = d                          ''�X�V���֌��݂̓��t
            End If
        Next s
    Next i
    
    Range(Cells(11, 12), Cells(9010, 12)) = koushin       ''�X�V���y�[�X�g
    
End Sub

Sub �X�V��_�R�[�v()
    Call �ی�.�S�ی����
    
    Worksheets("�����").Activate
    
    With ActiveSheet
        .Range("A10").Select
        If .FilterMode Then .ShowAllData
    End With
    
    Dim cn_bool As Variant, koushin As Variant
    Dim i As Long, s As Long, d As Date
    d = Date
    cn_bool = Range(Cells(11, 27), Cells(9010, 27))     ''�ܖ������ύX�O
    koushin = Range(Cells(11, 12), Cells(9010, 12))        ''�X�V��
    
    For i = 1 To 8999                                      ''��r����
            If cn_bool(i, 1) = 1 Or cn_bool(i, 1) = 2 Then
                koushin(i, 1) = d                          ''�X�V���֌��݂̓��t
            End If
    Next i
    Range(Cells(11, 12), Cells(9010, 12)) = koushin       ''�X�V���y�[�X�g
    MsgBox "�X�V���̍X�V���� �R�[�v"
End Sub
Sub �X�V��_IY()
    Call �ی�.�S�ی����
    
    Worksheets("�����").Activate
    
    With ActiveSheet
        .Range("A10").Select
        If .FilterMode Then .ShowAllData
    End With
    
    Dim cn_bool As Variant, koushin As Variant
    Dim i As Long, s As Long, d As Date
    d = Date
    cn_bool = Range(Cells(11, 31), Cells(9010, 31))     ''�ܖ������ύX�O
    koushin = Range(Cells(11, 12), Cells(9010, 12))        ''�X�V��
    
    For i = 1 To 8999                                      ''��r����
            If cn_bool(i, 1) = 1 Or cn_bool(i, 1) = 2 Then
                koushin(i, 1) = d                          ''�X�V���֌��݂̓��t
            End If
    Next i
    Range(Cells(11, 12), Cells(9010, 12)) = koushin       ''�X�V���y�[�X�g
    MsgBox "�X�V���̍X�V���� IY"
End Sub

Sub �X�V��_�����ޗ�()
    Call �ی�.�S�ی����
    
    Worksheets("�����").Activate
    
    With ActiveSheet
        .Range("A10").Select
        If .FilterMode Then .ShowAllData
    End With
    
    Dim cn_bool As Variant, koushin As Variant
    Dim i As Long, s As Long, d As Date
    d = Date
    cn_bool = Range(Cells(11, 26), Cells(9010, 26))     ''�ܖ������ύX�O
    koushin = Range(Cells(11, 12), Cells(9010, 12))        ''�X�V��
    
    For i = 1 To 8999                                      ''��r����
            If cn_bool(i, 1) = 1 Or cn_bool(i, 1) = 2 Then
                koushin(i, 1) = d                          ''�X�V���֌��݂̓��t
            End If
    Next i
    Range(Cells(11, 12), Cells(9010, 12)) = koushin       ''�X�V���y�[�X�g
    MsgBox "�X�V���̍X�V���� �����ޗ�"
End Sub
Sub �X�V��_����_�y()
    Call �ی�.�S�ی����
    
    Worksheets("�����").Activate
    
    With ActiveSheet
        .Range("A10").Select
        If .FilterMode Then .ShowAllData
    End With
    
    Dim cn_bool As Variant, koushin As Variant
    Dim i As Long, s As Long, d As Date
    d = Date
    cn_bool = Range(Cells(11, 45), Cells(9010, 45))     ''�ܖ������ύX�O
    koushin = Range(Cells(11, 12), Cells(9010, 12))        ''�X�V��
    
    For i = 1 To 8999                                      ''��r����
            If cn_bool(i, 1) = 1 Or cn_bool(i, 1) = 2 Then
                koushin(i, 1) = d                          ''�X�V���֌��݂̓��t
            End If
    Next i
    Range(Cells(11, 12), Cells(9010, 12)) = koushin       ''�X�V���y�[�X�g
    MsgBox "�X�V���̍X�V���� ����_�y"
End Sub
Sub �X�V��_����_��()
    Call �ی�.�S�ی����
    
    Worksheets("�����").Activate
    
    With ActiveSheet
        .Range("A10").Select
        If .FilterMode Then .ShowAllData
    End With
    
    Dim cn_bool As Variant, koushin As Variant
    Dim i As Long, s As Long, d As Date
    d = Date
    cn_bool = Range(Cells(11, 44), Cells(9010, 44))     ''�ܖ������ύX�O
    koushin = Range(Cells(11, 12), Cells(9010, 12))        ''�X�V��
    
    For i = 1 To 8999                                      ''��r����
            If cn_bool(i, 1) = 1 Or cn_bool(i, 1) = 2 Then
                koushin(i, 1) = d                          ''�X�V���֌��݂̓��t
            End If
    Next i
    Range(Cells(11, 12), Cells(9010, 12)) = koushin       ''�X�V���y�[�X�g
    MsgBox "�X�V���̍X�V���� ����_��"
End Sub

Sub �X�V��_����_��()
    Call �ی�.�S�ی����
    
    Worksheets("�����").Activate
    
    With ActiveSheet
        .Range("A10").Select
        If .FilterMode Then .ShowAllData
    End With
    
    Dim cn_bool As Variant, koushin As Variant
    Dim i As Long, s As Long, d As Date
    d = Date
    cn_bool = Range(Cells(11, 43), Cells(9010, 43))     ''�ܖ������ύX�O
    koushin = Range(Cells(11, 12), Cells(9010, 12))        ''�X�V��
    
    For i = 1 To 8999                                      ''��r����
            If cn_bool(i, 1) = 1 Or cn_bool(i, 1) = 2 Then
                koushin(i, 1) = d                          ''�X�V���֌��݂̓��t
            End If
    Next i
    Range(Cells(11, 12), Cells(9010, 12)) = koushin       ''�X�V���y�[�X�g
    MsgBox "�X�V���̍X�V���� ����_��"
End Sub

Sub �X�V��_����_��()
    Call �ی�.�S�ی����
    
    Worksheets("�����").Activate
    
    With ActiveSheet
        .Range("A10").Select
        If .FilterMode Then .ShowAllData
    End With
    
    Dim cn_bool As Variant, koushin As Variant
    Dim i As Long, s As Long, d As Date
    d = Date
    cn_bool = Range(Cells(11, 42), Cells(9010, 42))     ''�ܖ������ύX�O
    koushin = Range(Cells(11, 12), Cells(9010, 12))        ''�X�V��
    
    For i = 1 To 8999                                      ''��r����
            If cn_bool(i, 1) = 1 Or cn_bool(i, 1) = 2 Then
                koushin(i, 1) = d                          ''�X�V���֌��݂̓��t
            End If
    Next i
    Range(Cells(11, 12), Cells(9010, 12)) = koushin       ''�X�V���y�[�X�g
    MsgBox "�X�V���̍X�V���� ����_��"
End Sub

Sub �X�V��_����_��()
    Call �ی�.�S�ی����
    
    Worksheets("�����").Activate
    
    With ActiveSheet
        .Range("A10").Select
        If .FilterMode Then .ShowAllData
    End With
    
    Dim cn_bool As Variant, koushin As Variant
    Dim i As Long, s As Long, d As Date
    d = Date
    cn_bool = Range(Cells(11, 41), Cells(9010, 41))     ''�ܖ������ύX�O
    koushin = Range(Cells(11, 12), Cells(9010, 12))        ''�X�V��
    
    For i = 1 To 8999                                      ''��r����
            If cn_bool(i, 1) = 1 Or cn_bool(i, 1) = 2 Then
                koushin(i, 1) = d                          ''�X�V���֌��݂̓��t
            End If
    Next i
    Range(Cells(11, 12), Cells(9010, 12)) = koushin       ''�X�V���y�[�X�g
    MsgBox "�X�V���̍X�V���� ����_��"
End Sub

Sub �X�V��_����_��()
    Call �ی�.�S�ی����
    
    Worksheets("�����").Activate
    
    With ActiveSheet
        .Range("A10").Select
        If .FilterMode Then .ShowAllData
    End With
    
    Dim cn_bool As Variant, koushin As Variant
    Dim i As Long, s As Long, d As Date
    d = Date
    cn_bool = Range(Cells(11, 40), Cells(9010, 40))     ''�ܖ������ύX�O
    koushin = Range(Cells(11, 12), Cells(9010, 12))        ''�X�V��
    
    For i = 1 To 8999                                      ''��r����
            If cn_bool(i, 1) = 1 Or cn_bool(i, 1) = 2 Then
                koushin(i, 1) = d                          ''�X�V���֌��݂̓��t
            End If
    Next i
    Range(Cells(11, 12), Cells(9010, 12)) = koushin       ''�X�V���y�[�X�g
    MsgBox "�X�V���̍X�V���� ����_��"
End Sub

Sub �X�V��_����_��()
    Call �ی�.�S�ی����
    
    Worksheets("�����").Activate
    
    With ActiveSheet
        .Range("A10").Select
        If .FilterMode Then .ShowAllData
    End With
    
    Dim cn_bool As Variant, koushin As Variant
    Dim i As Long, s As Long, d As Date
    d = Date
    cn_bool = Range(Cells(11, 39), Cells(9010, 39))     ''�ܖ������ύX�O
    koushin = Range(Cells(11, 12), Cells(9010, 12))        ''�X�V��
    
    For i = 1 To 8999                                      ''��r����
            If cn_bool(i, 1) = 1 Or cn_bool(i, 1) = 2 Then
                koushin(i, 1) = d                          ''�X�V���֌��݂̓��t
            End If
    Next i
    Range(Cells(11, 12), Cells(9010, 12)) = koushin       ''�X�V���y�[�X�g
    MsgBox "�X�V���̍X�V���� ����_��"
End Sub

