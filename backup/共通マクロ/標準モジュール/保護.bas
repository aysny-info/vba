Attribute VB_Name = "�ی�"

Sub �ی�()
    ActiveSheet.Protect Contents:=True, DrawingObjects:=False, UserInterfaceOnly:=True, AllowFormattingCells:=True
End Sub

Sub �ی����()
    ActiveSheet.Unprotect
End Sub

Sub �ی�_����()
    For ws_num = 1 To (Sheets("���v���z").Index - 1)
        Worksheets(ws_num).Protect Contents:=True, DrawingObjects:=False, UserInterfaceOnly:=True, AllowFormattingCells:=True
    Next ws_num
    
End Sub

Sub �ی�_�S����()
    Dim sh As Object
    On Error Resume Next
    For Each sh In Sheets
    sh.Unprotect
    Next sh
    
    Worksheets(1).Activate
End Sub

Sub �V�[�g�����I������()
    ActiveWindow.SelectedSheets(1).Select
End Sub

Sub �ی샍�b�N�S�V�[�g�z��()

    '�V�[�g�z�@���v���z�V�[�g���Ђ���
    For ws_num = 1 To (Sheets("���v���z").Index - 1)

        '�B��V�[�g�͔�΂�
        If Not Worksheets(ws_num).Visible Then
            GoTo CONTINUE
        End If

        '�u�I���\�v�u�����W�J�v�V�[�g��΂�
        If Worksheets(ws_num).Name = "�I���\" Or Worksheets(ws_num).Name = "�����W�J" Then
            GoTo CONTINUE
        End If

        '�\���V�[�g�A�N�e�B�u
        Worksheets(ws_num).Activate
        
        '���s
        Call �ی샍�b�N����obj
CONTINUE:
    Next ws_num
End Sub

Sub �ی샍�b�N����obj()
    '�S�Z���ی�
     Cells.Locked = True

    '�ŏI�s�擾
    Dim LastRow As Long
    LastRow = Cells(Rows.Count, 5).End(xlUp).Row

    'E�񌟍��p
    serch_cell = Range(Cells(1, 5), Cells(LastRow, 9)).Formula
    
    '�P�`�R�P��
    Dim tikan_cell As Range
    Set tikan_cell = Range(Cells(1, 8), Cells(LastRow, 38))
    
    '�Y���Z�������b�N����
    For s = 1 To LastRow
        If serch_cell(s, 1) = "���א�" Or _
            serch_cell(s, 1) = "���v���א�" Or _
            serch_cell(s, 1) = "�����R�[�q�[" Or _
            serch_cell(s, 1) = "�T�|�[�g" Or _
            serch_cell(s, 1) = "���l���}" Or _
            serch_cell(s, 1) = "�ԕi��" Or _
            serch_cell(s, 1) = "�a��" Or _
            serch_cell(s, 1) = "�߂�" _
            Then
            
            Debug.Print ActiveSheet.Name & "___�ی샍�b�N��������Z���I��__" & s & "�s"
            
            For n = 1 To 31
                tikan_cell(s, n).Locked = False
            Next n
            
        End If
        
        If serch_cell(s, 1) = "����" Or _
            serch_cell(s, 1) = "����1" Or _
            serch_cell(s, 1) = "����2" _
            Then
                
            '�����������Ă��邩�ǂ���
            If Mid(serch_cell(s, 4), 1, 1) = "=" Then
                Debug.Print ActiveSheet.Name & "___�ی샍�b�N��������Z���I�����Ȃ�__" & s & "�s"
            Else
                Debug.Print ActiveSheet.Name & "___�ی샍�b�N��������Z���I��__����__" & s & "�s"
                
                For n = 1 To 31
                    tikan_cell(s, n).Locked = False
                Next n
            End If
        End If
    Next s
    
End Sub
