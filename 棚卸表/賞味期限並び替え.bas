Attribute VB_Name = "�ܖ��������ёւ�"
'''''''''''''''''''''''''''''''        �u������v�֘A         ''''''''''''''''''
Sub �ܖ��������ёւ�()
    Call �ی�.�S�ی����

    Worksheets("�����").Activate
    
    With ActiveSheet
        .Range("A10").Select
        If .FilterMode Then .ShowAllData
    End With
    
    '�\�[�g
    Dim y As Long, i As Long, j As Long, v As Long, swap As Variant, B As Variant, C As Variant
    B = Range(Cells(11, 13), Cells(9010, 18))

    For y = 1 To 8999 ''�\�[�g�J�n
        For i = 1 To 6
            For j = 1 To 6
                If B(y, i) < B(y, j) Then
                    swap = B(y, i)
                    B(y, i) = B(y, j)
                    B(y, j) = swap
                End If
            Next j
        Next i
    Next y

    For y = 1 To 8999 ''���l��
        For v = 1 To 5
         For i = 1 To 5
             If B(y, i) = "" Then
                 For j = i To 5
                     B(y, j) = B(y, j + 1)
                     B(y, j + 1) = ""
                 Next j
             End If
        Next i
       Next v
   Next y

    Range(Cells(11, 13), Cells(9010, 18)) = B
    
    'Call �X�V��.�X�V��_����
    Call �ی�.�����ی�
    MsgBox "�ܖ��������ёւ�����"
End Sub


