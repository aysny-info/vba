Sub Error_check()
   '�ۑ������Ƃ��̃V�[�g
    first_name = ActiveSheet.Name
    
    If Workbooks(1).Name Like "*�������i*" Then
        nuke = kowake_check()
    Else
        nuke = normal()
    End If
    
     '�A�N�e�B�u�V�[�g�����Ƃɖ߂�
    ActiveWorkbook.Sheets(first_name).Activate

    If nuke = 1 Then
        MsgBox "�����������Ă��܂��B�m�F���Ă��������B"
    End If
End Sub
Function normal() As Variant
    ActiveWorkbook.Sheets("���v���z").Activate
    '�ŏI�s�擾
        Dim LastRow As Long
        LastRow = Cells(Rows.Count, 9).End(xlUp).Row
        Debug.Print LastRow
        '������z��֊i�[     I��
        search_cell = Range(Cells(1, 9), Cells(LastRow, 9))

        Dim nuke As Integer
        nuke = 0
        For s = 1 To LastRow
               
                    If search_cell(s, 1) = "�G���[" Then
                          nuke = 1
                    End If
              
        Next s
        normal = nuke
End Function

Function kowake_check() As Variant
    ActiveWorkbook.Sheets("�������i").Activate
    If Range("BB6") = "�G���[" Then
        nuke = 1
    End If
    kowake_check = nuke
End Function
