Attribute VB_Name = "�����u��"
Sub �u��main()
    '�Z�b�e�B���O �v�v�v�v�v
    Set dateList = ���t�N���X()
    Set fileList = �t�@�C���N���X()
    Set colorList = �J���[�N���X()

    Call ���s�O�`�F�b�N(dateList, fileList)
    Call �S�V�[�g�z��(dateList, fileList, colorList)
    Call �����W�J�̒u��(dateList, fileList) '�d���@�|���݌ɕ\��p
    Call ���v���z�V�[�gB2�����֕ύX(dateList, fileList)
End Sub

Sub ���s�O�`�F�b�N(dateList As Variant, fileList As Variant)
    '�t�@�C�����`�F�b�N
'    If fileList.bool_filename Then
'        Debug.Print "OK"
'    Else
'        Debug.Print fileList.bool_filename
'
'        MsgBox ("�t�@�C����orVBA�̃t�@�C�����`�F�b�N�����������ł��B" & vbCrLf & fileList.this_filename & vbCrLf & fileList.checkFilename)
'        End
'    End If

    '���v���z�V�[�gB2�Z���`�F�b�N
    Dim boool As Long
    If fileList.mybook_month = Month(dateList.date_next) Then
        boool = MsgBox("�����݌ɕ\�쐬�̂��߁u�u���A�폜�v�@���s���܂����A��낵���ł����H", vbYesNo + vbQuestion)
        If boool = vbYes Then
            Exit Sub
        Else
            End
        End If
    Else
        MsgBox ("�@�t�@�C�����̌�" & vbCrLf & "�A�u���v���z�v�V�[�g�uB2�v�Z��" & vbCrLf & "�̊֌W�����������ł��B" & vbCrLf & vbCrLf & "��������" & vbCrLf & "�@�݌Ɂw��ށx_2020.8.xlsm" & vbCrLf & "�A2020/7/1")
        End
    End If
End Sub

Sub �S�V�[�g�z��(dateList As Variant, fileList As Variant, colorList As Variant)
    '�J���[�w��p�z��ϐ��̐錾
    Dim colorGenArray() As Variant
    Dim colorSheetArray() As Variant
        '���
    colorGenArray = colorList.colorGenArray
    colorSheetArray = colorList.colorSheetArray

    'H-AN��ĕ\�� �`���א����� �v�v�v�v�v
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
        Call H_AL��u���폜(dateList, fileList)
        Call G��u��(dateList.G_last_date, dateList.G_now_date)
        Call ���א��̐F(dateList, fileList)
        Call �a���̐F(dateList, fileList)
        Call �߂��̐F(dateList, fileList)

        If ActiveSheet.Name = "�R�[�v_�p�b�N" Then
            Call �T�|�[�g�̐F(dateList, fileList)
            Call ���l���}�̐F(dateList, fileList)
        End If

        If ActiveSheet.Name = "�_�c���Y��" Then
            Call �����R�[�q�[�̐F(dateList, fileList)
        End If
        
        'Filter()�֐��ɂ��v�f������
        varResult = Filter(colorSheetArray, ActiveSheet.Name)
        If UBound(varResult) <> -1 Then
            Debug.Print "�J���[�w��V�[�g__" & ActiveSheet.Name
            Call �J���[�w��(colorGenArray)
        End If
        
CONTINUE:
    Next ws_num
End Sub

Sub �����W�J�̒u��(dateList As Variant, fileList As Variant)
        If fileList.this_filename <> "�݌Ɂw�|���x_" Then
            Debug.Print "�����W�J�̒u���͂��Ȃ�"
            Exit Sub
        End If
        Debug.Print "�����W�J�̒u���J�n"

        ActiveWorkbook.Sheets("�����W�J").Activate
        '�ŏI�s�擾
        Dim LastRow As Long
        LastRow = Cells(Rows.Count, 2).End(xlUp).Row
        Debug.Print LastRow
        '������z��֊i�[     F�񂩂�AJ��
        tikan_cell = Range(Cells(2, 6), Cells(LastRow, 36)).Formula

        '�u��
        For s = 1 To (LastRow - 1)
                For n = 1 To 31
                    tikan_cell(s, n) = Replace(tikan_cell(s, n), dateList.H_AL_nowMonth, dateList.H_AL_nextMonth)
                Next n
        Next s

        '�z����V�[�g�� �y�[�X�g
        Range(Cells(2, 6), Cells(LastRow, 36)).Formula = tikan_cell
End Sub

Sub H_AL��u���폜(dateList As Variant, fileList As Variant)
        '�ŏI�s�擾
        Dim LastRow1, LastRow2, LastRow3, LastRow4, LastRow5, LastRow As Long
        LastRow1 = Cells(Rows.Count, 5).End(xlUp).Row
        LastRow2 = Cells(Rows.Count, 6).End(xlUp).Row
        LastRow3 = Cells(Rows.Count, 7).End(xlUp).Row
        LastRow4 = Cells(Rows.Count, 8).End(xlUp).Row
        LastRow5 = Cells(Rows.Count, 38).End(xlUp).Row

        Dim arr As Variant
        arr = Array(LastRow1, LastRow2, LastRow3, LastRow4, LastRow5)
        LastRow = WorksheetFunction.Max(arr)

        '������z��֊i�[     E�񂩂�AL��
        tikan_cell = Range(Cells(1, 5), Cells(LastRow, 38)).Formula

        '������z��֊i�[     AO�񂩂�BS��
        If fileList.this_filename = "�݌Ɂw��ށx_" Then
            tikan_222 = Range(Cells(1, 41), Cells(LastRow, 71)).Formula
        End If

        'E�������
        Debug.Print ActiveSheet.Name & " E��ŏI�s��" & LastRow
        For s = 1 To LastRow

        '���א��̍폜
        If ActiveSheet.Name <> "�����R�[�q�[�t�[�Y��" Then
            '������1���`�V���ɓ����Ă��邩�ǂ���
            If Mid(tikan_cell(s, 4), 1, 1) = "=" And Mid(tikan_cell(s, 5), 1, 1) = "=" And Mid(tikan_cell(s, 6), 1, 1) = "=" And Mid(tikan_cell(s, 7), 1, 1) = "=" And Mid(tikan_cell(s, 8), 1, 1) = "=" And Mid(tikan_cell(s, 9), 1, 1) = "=" And Mid(tikan_cell(s, 10), 1, 1) = "=" Then
                Debug.Print ActiveSheet.Name & s & "�s " & "���א��ɐ�������̂ō폜���Ȃ�"
            Else
                If tikan_cell(s, 1) = "���א�" Or tikan_cell(s, 1) = "���v���א�" Or tikan_cell(s, 1) = "�����R�[�q�[" Then
                    Debug.Print ActiveSheet.Name & s & "�s " & "���א��̍폜"
                    For n = 4 To 34
                        tikan_cell(s, n) = ""
                    Next n
                End If
            End If
        End If

            '�u�R�[�v_�p�b�N�v�V�[�g�̓��א��폜
            If ActiveSheet.Name = "�R�[�v_�p�b�N" Then
                '�T�|�[�g�̍폜
                If tikan_cell(s, 1) = "�T�|�[�g" Then
                    Debug.Print "�T�|�[�g���א�"
                    For n = 4 To 34
                        tikan_cell(s, n) = ""
                    Next n
                End If

                '���l���}�̍폜
                If tikan_cell(s, 1) = "���l���}" Then
                    Debug.Print "���l���}���א�"
                    For n = 4 To 34
                        tikan_cell(s, n) = ""
                    Next n
                End If
            End If

            '�����̍폜
            If tikan_cell(s, 1) = "����" Then
                '�����������Ă��邩�ǂ���
                If Mid(tikan_cell(s, 4), 1, 1) = "=" Then
                    Debug.Print ActiveSheet.Name & s & "�s " & "�����ɐ�������̂ō폜���Ȃ�"
                Else
                    Debug.Print ActiveSheet.Name & s & "�s " & "�����폜"
                    For n = 4 To 34
                        tikan_cell(s, n) = ""
                    Next n
                End If
            End If

            '�_���{�[���A�o�א��̍폜
            If ActiveSheet.Name = "���R���" Then
                If tikan_cell(s, 1) = "�o�א�(�����)" Then
                    '�����������Ă��邩�ǂ���
                    If Mid(tikan_cell(s, 4), 1, 1) = "=" Then
                        Debug.Print ActiveSheet.Name & s & "�s " & "�o�א�(�����)�ɐ�������̂ō폜���Ȃ�"
                    Else
                        Debug.Print ActiveSheet.Name & s & "�s " & "�o�א�(�����)�폜"
                        For n = 4 To 34
                            tikan_cell(s, n) = ""
                        Next n
                    End If
                End If
            End If


            '�ԕi���̍폜
            If tikan_cell(s, 1) = "�ԕi��" Then
                Debug.Print ActiveSheet.Name & s & "�s " & "�ԕi��"
                For n = 4 To 34
                    tikan_cell(s, n) = ""
                Next n
            End If


            '�a���̍폜
            If tikan_cell(s, 1) = "�a��" Then
                Debug.Print ActiveSheet.Name & s & "�s " & "�a��"
                For n = 4 To 34
                    tikan_cell(s, n) = ""
                Next n
            End If


            '�߂��̍폜
            If tikan_cell(s, 1) = "�߂�" Then
                Debug.Print ActiveSheet.Name & s & "�s " & "�߂�"
                For n = 4 To 34
                    tikan_cell(s, n) = ""
                Next n
            End If

            '�o�א��̒u��
'            If tikan_cell(s, 1) = "�o�א�" Then
'                Debug.Print ActiveSheet.Name & s &  "�o�א�"
'                For n = 4 To 34
'                    tikan_cell(s, n) = Replace(tikan_cell(s, n), dateList.H_AL_nowMonth, dateList.H_AL_nextMonth)
'                Next n
'            End If
            '�v�搔�̒u��
'            If tikan_cell(s, 1) = "�v�搔" Then
'                Debug.Print ActiveSheet.Name & s &  "�v�搔"
'                For n = 4 To 34
'                    tikan_cell(s, n) = Replace(tikan_cell(s, n), dateList.H_AL_nowMonth, dateList.H_AL_nextMonth)
'                Next n
'            End If
            '�����̒u��
'            If tikan_cell(s, 1) = "����" Then
'                Debug.Print ActiveSheet.Name & s &  "����"
'                For n = 4 To 34
'                    tikan_cell(s, n) = Replace(tikan_cell(s, n), dateList.H_AL_nowMonth, dateList.H_AL_nextMonth)
'                Next n
'            End If
            '�T���v���̒u��
'            If tikan_cell(s, 1) = "�T���v��" Then
'                Debug.Print ActiveSheet.Name & s &  "�T���v��"
'                For n = 4 To 34
'                    tikan_cell(s, n) = Replace(tikan_cell(s, n), dateList.H_AL_nowMonth, dateList.H_AL_nextMonth)
'                Next n
'            End If
'            '�d�|�ʂ̒u��
'            If tikan_cell(s, 1) = "�d�|��" Then
'                Debug.Print ActiveSheet.Name & s & "�s " & "�d�|��"
'                For n = 4 To 34
'                    tikan_cell(s, n) = Replace(tikan_cell(s, n), dateList.G_now_date, dateList.G_next_date)
'                Next n
'            End If

            '**************���̒u��**************
'            If tikan_cell(s, 1) = "�䗦&�K�v��" Or tikan_cell(s, 1) = "�䗦" Then
'                Debug.Print ActiveSheet.Name & s & "�s " &  "�䗦&�K�v��"
'                For k = s To LastRow
'                    For n = 4 To 34
'                        tikan_cell(k, n) = Replace(tikan_cell(k, n), dateList.H_AL_nowMonth, dateList.H_AL_nextMonth)
'                    Next n
'                Next k
'                s = LastRow
'            End If
            '**************���̒u��**************
            '�������̒u��
'            If tikan_cell(s, 3) = "������" Then
'                Debug.Print ActiveSheet.Name & s & "�s " &  "������"
'                For n = 4 To 34
'                    tikan_cell(s, n) = Replace(tikan_cell(s, n), dateList.kowake_now_date, dateList.kowake_next_date)
'                Next n
'            End If

        Next s

        '**************H�`AL�̒u��**************
        Debug.Print ActiveSheet.Name & "H�`AL�̒u��"
        For k = 1 To LastRow
            For n = 4 To 34
                tikan_cell(k, n) = Replace(tikan_cell(k, n), dateList.H_AL_nowMonth, dateList.H_AL_nextMonth)
                tikan_cell(k, n) = Replace(tikan_cell(k, n), dateList.kowake_now_date, dateList.kowake_next_date)
                tikan_cell(k, n) = Replace(tikan_cell(k, n), dateList.G_now_date, dateList.G_next_date)
            Next n
        Next k

        '**************AO�`BS�̒u��**************
        If fileList.this_filename = "�݌Ɂw��ށx_" Then
            Debug.Print ActiveSheet.Name & "AO�`BS�̒u��"
            For k = 1 To LastRow
                For n = 1 To 31
                    tikan_222(k, n) = Replace(tikan_222(k, n), dateList.H_AL_nextMonth, dateList.H_AL_Month_after_next)
                Next n
            Next k
        End If

        '�z����V�[�g�� �y�[�X�g
        On Error Resume Next
        Range(Cells(1, 5), Cells(LastRow, 38)).Formula = tikan_cell

        '���AO�`BS�z����V�[�g�� �y�[�X�g
        If fileList.this_filename = "�݌Ɂw��ށx_" Then
            On Error Resume Next
            Range(Cells(1, 41), Cells(LastRow, 71)).Formula = tikan_222
        End If

        '�R�����g�폜
        Range("H:AL").ClearComments

        '���AO�`BS�R�����g�폜
        If fileList.this_filename = "�݌Ɂw��ށx_" Then
            Range("AO:BS").ClearComments
        End If
End Sub

Sub G��u��(ByVal now_tsuki As String, ByVal next_tsuki As String)
    '�ŏI�s�擾
    Dim LastRow As Long
    LastRow = Cells(Rows.Count, 7).End(xlUp).Row
    'G�� tikan_cell
    tikan_cell = Range(Cells(1, 7), Cells(LastRow, 7)).Formula
    '�u��
    For i = 1 To LastRow
        tikan_cell(i, 1) = Replace(tikan_cell(i, 1), now_tsuki, next_tsuki)
    Next i

    Range(Cells(1, 7), Cells(LastRow, 7)).Formula = tikan_cell
End Sub

'Sub ���v���z�V�[�gB2�����֕ύX(dateList As Variant, fileList As Variant)
'    ActiveWorkbook.Sheets("���v���z").Activate
'    Range("B2").Formula = dateList.date_next '���v���z�V�[�gB2�����֕ύX
'End Sub

Sub ���א��̐F(dateList As Variant, fileList As Variant)
    '�ŏI�s�擾
    Dim LastRow As Long
    LastRow = Cells(Rows.Count, 5).End(xlUp).Row

    '������z��֊i�[     E�񂩂�AL�� �u���א��v�����p
    tikan_cell = Range(Cells(1, 5), Cells(LastRow, 5)).Formula

    '�Z����select���R�s�[
    For s = 1 To LastRow
        If tikan_cell(s, 1) = "���א�" Then
            Cells(s, 39).Select 'select�Z���̃��Z�b�g
            Cells(s, 39).Copy   '���א��̍��v�Z��(AM4)�R�s�[ (���R��ReFAX���F�ύX���Ȃ��Z��������j
            Exit For
        End If
    Next s

    '���א�select
    For s = 1 To LastRow
        '���א��̍폜
        If tikan_cell(s, 1) = "���א�" Then
            Debug.Print ActiveSheet.Name & "���א�_�F���f�t�H���g��"
            For n = 4 To 34
                Union(Selection, Range(Cells(s, 8), Cells(s, 39))).Select
            Next n
        End If
    Next s

    '�y�[�X�g
    'ActiveSheet.Paste
        Selection.PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, _
        SkipBlanks:=False, Transpose:=False
End Sub

Sub �����R�[�q�[�̐F(dateList As Variant, fileList As Variant)
    '�ŏI�s�擾
    Dim LastRow As Long
    LastRow = Cells(Rows.Count, 5).End(xlUp).Row

    '������z��֊i�[     E�񂩂�AL�� �u�����R�[�q�[�v�����p
    tikan_cell = Range(Cells(1, 5), Cells(LastRow, 5)).Formula

    '�Z����select���R�s�[
    For s = 1 To LastRow
        If tikan_cell(s, 1) = "�����R�[�q�[" Then
            Cells(s, 39).Select 'select�Z���̃��Z�b�g
            Cells(s, 39).Copy   '�����R�[�q�[�̍��v�Z��(AM4)�R�s�[ (���R��ReFAX���F�ύX���Ȃ��Z��������j
            Exit For
        End If
    Next s

    '�����R�[�q�[select
    For s = 1 To LastRow
        '�����R�[�q�[�̍폜
        If tikan_cell(s, 1) = "�����R�[�q�[" Then
            Debug.Print ActiveSheet.Name & "�����R�[�q�[_�F���f�t�H���g��"
            For n = 4 To 34
                Union(Selection, Range(Cells(s, 8), Cells(s, 39))).Select
            Next n
        End If
    Next s

    '�y�[�X�g
    'ActiveSheet.Paste
        Selection.PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, _
        SkipBlanks:=False, Transpose:=False
End Sub

Sub �T�|�[�g�̐F(dateList As Variant, fileList As Variant)
    '�ŏI�s�擾
    Dim LastRow As Long
    LastRow = Cells(Rows.Count, 5).End(xlUp).Row

    '������z��֊i�[     E�񂩂�AL�� �u���א��v�����p
    tikan_cell = Range(Cells(1, 5), Cells(LastRow, 5)).Formula

    '�Z����select���R�s�[
    For s = 1 To LastRow
        If tikan_cell(s, 1) = "�T�|�[�g" Then
            Cells(s, 39).Select 'select�Z���̃��Z�b�g
            Cells(s, 39).Copy   '���א��̍��v�Z��(AM4)�R�s�[ (���R��ReFAX���F�ύX���Ȃ��Z��������j
            Exit For
        End If
    Next s

    '���א�select
    For s = 1 To LastRow
        '���א��̍폜
        If tikan_cell(s, 1) = "�T�|�[�g" Then
            Debug.Print ActiveSheet.Name & "���א�_�F���f�t�H���g��"
            For n = 4 To 34
                Union(Selection, Range(Cells(s, 8), Cells(s, 39))).Select
            Next n
        End If
    Next s

    '�y�[�X�g
    'ActiveSheet.Paste
        Selection.PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, _
        SkipBlanks:=False, Transpose:=False
End Sub

Sub ���l���}�̐F(dateList As Variant, fileList As Variant)
    '�ŏI�s�擾
    Dim LastRow As Long
    LastRow = Cells(Rows.Count, 5).End(xlUp).Row

    '������z��֊i�[     E�񂩂�AL�� �u���א��v�����p
    tikan_cell = Range(Cells(1, 5), Cells(LastRow, 5)).Formula

    '�Z����select���R�s�[
    For s = 1 To LastRow
        If tikan_cell(s, 1) = "���l���}" Then
            Cells(s - 1, 39).Select '(s -1)�̓T�|�[�g�̐F �Z��[AM5]
            Cells(s - 1, 39).Copy '(AM5)�R�s�[ (���R��ReFAX���F�ύX���Ȃ��Z��������j
            Exit For
        End If
    Next s

    '���א�select
    For s = 1 To LastRow
        '���א��̍폜
        If tikan_cell(s, 1) = "���l���}" Then
            Debug.Print ActiveSheet.Name & "���א�_�F���f�t�H���g��"
            For n = 4 To 33
                Union(Selection, Range(Cells(s, 8), Cells(s, 38))).Select
            Next n
        End If
    Next s

    '�y�[�X�g
    'ActiveSheet.Paste
        Selection.PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, _
        SkipBlanks:=False, Transpose:=False
End Sub


Sub �a���̐F(dateList As Variant, fileList As Variant)
    '�ŏI�s�擾
    Dim LastRow As Long
    LastRow = Cells(Rows.Count, 5).End(xlUp).Row

    '������z��֊i�[     E�񂩂�AL�� �u�a���v�����p
    tikan_cell = Range(Cells(1, 5), Cells(LastRow, 5)).Formula

    '�Z����select���R�s�[
    For s = 1 To LastRow
        If tikan_cell(s, 1) = "�a��" Then
            Cells(s, 39).Select 'select�Z���̃��Z�b�g
            Cells(s, 39).Copy   '�߂��Z��(AM12�t��)�R�s�[ (���R�͐F�ύX���Ȃ��Z��������j
            Exit For
        End If
    Next s

    '���א�select
    For s = 1 To LastRow
        '���א��̍폜
        If tikan_cell(s, 1) = "�a��" Then
            Debug.Print ActiveSheet.Name & "�a��_�F���f�t�H���g��"
            For n = 4 To 34
                Union(Selection, Range(Cells(s, 8), Cells(s, 39))).Select
            Next n
        End If
    Next s

    '�y�[�X�g
    'ActiveSheet.Paste
        Selection.PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, _
        SkipBlanks:=False, Transpose:=False
End Sub

Sub �߂��̐F(dateList As Variant, fileList As Variant)
    '�ŏI�s�擾
    Dim LastRow As Long
    LastRow = Cells(Rows.Count, 5).End(xlUp).Row

    '������z��֊i�[     E�񂩂�AL�� �u�߂��v�����p
    tikan_cell = Range(Cells(1, 5), Cells(LastRow, 5)).Formula

    '�Z����select���R�s�[
    For s = 1 To LastRow
        If tikan_cell(s, 1) = "�߂�" Then
            Cells(s, 39).Select 'select�Z���̃��Z�b�g
            Cells(s, 39).Copy   '�߂��Z��(AM12�t��)�R�s�[ (���R�͐F�ύX���Ȃ��Z��������j
            Exit For
        End If
    Next s

    '���א�select
    For s = 1 To LastRow
        '���א��̍폜
        If tikan_cell(s, 1) = "�߂�" Then
            Debug.Print ActiveSheet.Name & "�߂�_�F���f�t�H���g��"
            For n = 4 To 34
                Union(Selection, Range(Cells(s, 8), Cells(s, 39))).Select
            Next n
        End If
    Next s

    '�y�[�X�g
    'ActiveSheet.Paste
        Selection.PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, _
        SkipBlanks:=False, Transpose:=False
End Sub

Sub �J���[�w��(colorGenArray As Variant)
    '�ŏI�s�擾
    Dim LastRow As Long
    Dim rgb_value As Long
    LastRow = Cells(Rows.Count, 3).End(xlUp).Row
    
    '������z��֊i�[     E�񂩂�AL��
    tikan_cell = Range(Cells(1, 3), Cells(LastRow, 5)).Formula
    
    'C�񂪂S�����@���@E��u���א��vor�u���v���א��v ���� �J���[�w�茴���R�[�h��v
    For s = 1 To LastRow
        If Len(tikan_cell(s, 1)) = 4 Then
            If tikan_cell(s, 3) = "���א�" Or tikan_cell(s, 3) = "���v���א�" Then
                'Filter()�֐��ɂ��v�f������
                varResult = Filter(colorGenArray, tikan_cell(s, 1))
                
                '�J���[�w�茴���R�[�h����v�����ꍇ�AAN��F�R�s�y
                If UBound(varResult) <> -1 Then
                    '�ȑO��2020.09.10
'                                Cells(s, 39).Select 'select�Z���̃��Z�b�g
'                                Cells(s, 39).Copy   '���א��̍��v�Z��(AM4)�R�s�[ (���R��ReFAX���F�ύX���Ȃ��Z��������j
'                                Range(Cells(s, 8), Cells(s, 39)).Select '�P���`�R�P���܂ŃZ���N�g
'
'                            '�y�[�X�g
'                                Selection.PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, _
'                                SkipBlanks:=False, Transpose:=False
                        rgb_value = Cells(s, 40).Interior.Color 'AN�̐F���擾
                        Range(Cells(s, 8), Cells(s, 39)).Interior.Color = rgb_value
                    
                    Debug.Print "�J���[�w��___" & tikan_cell(s, 1) & tikan_cell(s, 3)
                End If
            End If
        End If
    Next s

End Sub

'Sub �f�[�^�`�F�b�N(dateList As Variant, fileList As Variant)
'   Debug.Print "�udateList.date_now�v " & dateList.date_now
'   Debug.Print "�udateList.date_next�v " & dateList.date_next
'   Debug.Print "�udateList.date_last�v " & dateList.date_last
'   Debug.Print "�udateList.H_AL_nowMonth�v " & dateList.H_AL_nowMonth
'   Debug.Print "�udateList.H_AL_nextMonth�v " & dateList.H_AL_nextMonth
'   Debug.Print "�udateList.G_now_date�v " & dateList.G_now_date
'   Debug.Print "�udateList.G_last_date�v " & dateList.G_last_date
'   Debug.Print "�udateList.G_next_date�v " & dateList.G_next_date
'   Debug.Print "�udateList.kowake_now_date�v " & dateList.kowake_now_date
'   Debug.Print "�udateList.kowake_last_date�v " & dateList.kowake_last_date
'   Debug.Print "�udateList.kowake_next_date�v " & dateList.kowake_next_date
'
'   Debug.Print "�ufileList.this_filename�v " & fileList.this_filename
'   Debug.Print "�ufileList.mypath�v " & fileList.mypath
'   Debug.Print "�ufileList.mybook�v " & fileList.mybook
'   Debug.Print "�ufileList.mybook_month�v " & fileList.mybook_month
'   Debug.Print "�ufileList.mmfn�v " & fileList.mmfn
'   Debug.Print "�ufileList.fn�v " & fileList.fn
'End Sub

Public Function ���t�N���X() As dateClass
    '�錾
    Dim dateList  As dateClass
    Set dateList = New dateClass

    '�Z�b�g
    ActiveWorkbook.Sheets("���v���z").Activate
    dateList.date_now = Range("B2")    '���v���z�V�[�g��B2���t
    dateList.date_after_next = DateAdd("m", 2, DateSerial(Year(dateList.date_now), Month(dateList.date_now), 1)) '���X��date
    dateList.date_next = DateAdd("m", 1, DateSerial(Year(dateList.date_now), Month(dateList.date_now), 1))       '����date
    dateList.date_last = DateAdd("m", -1, DateSerial(Year(dateList.date_now), Month(dateList.date_now), 1))       '�挎date
    dateList.H_AL_nowMonth = Format(dateList.date_now, "m��")    '�����ux���v
    dateList.H_AL_nextMonth = Format(dateList.date_next, "m��") '�����ux���v
    dateList.H_AL_Month_after_next = Format(dateList.date_after_next, "m��") '���X���ux���v
    dateList.G_now_date = Year(dateList.date_now) & "." & Month(dateList.date_now) '��u2020.7�v
    dateList.G_last_date = Year(dateList.date_last) & "." & Month(dateList.date_last) '��u2020.6�v
    dateList.G_next_date = Year(dateList.date_next) & "." & Month(dateList.date_next) '��u2020.8�v
    dateList.kowake_now_date = Year(dateList.date_now) & "." & Right("0" & Month(dateList.date_now), 2) '��u2020.07�v
    dateList.kowake_last_date = Year(dateList.date_last) & "." & Right("0" & Month(dateList.date_last), 2) '��u2020.06�v
    dateList.kowake_next_date = Year(dateList.date_next) & "." & Right("0" & Month(dateList.date_next), 2) '��u2020.08�v

    Set ���t�N���X = dateList
End Function

Public Function �t�@�C���N���X() As fileClass
    '�錾
    Dim fileList  As fileClass
    Set fileList = New fileClass
    '�Z�b�g
    fileList.mypath = ThisWorkbook.Path
    fileList.mybook = ThisWorkbook.Name
    fileList.mmfn = fileList.mypath & "\" & fileList.mybook
    'fileList.checkFilename = "�݌Ɂw�����ޗ��x_"

    '�t�@�C�����̌��𒊏o
    ActiveWorkbook.Sheets("���v���z").Activate
    date_now = Range("B2")    '���v���z�V�[�g��B2���t
    date_next = DateAdd("m", 1, DateSerial(Year(date_now), Month(date_now), 1))       '����date
'    If Month(date_now) = 12 Then '12���̏ꍇ
'        s = InStr(fileList.mybook, Year(date_next) & ".") + 5  '�u2020.�v�͉������ڂ���n�܂邩
'    Else                        '1�`11���̏ꍇ
'        s = InStr(fileList.mybook, Year(date_now) & ".") + 5  '�u2020.�v�͉������ڂ���n�܂邩
'    End If
    s = InStr(fileList.mybook, Year(date_now) & ".") + 5  '�u2020.�v�͉������ڂ���n�܂邩
    
    
    l = Len(fileList.mybook)
    str2 = Mid(fileList.mybook, s, l)   '�� �u8.xlsm�v
    s2 = InStr(str2, ".")
    l2 = Len(str2)

    Fname = InStr(fileList.mybook, Year(date_now) & ".") '�u2020.�v�͉������ڂ���n�܂邩
'    If Month(date_now) = 12 Then '12��
'        Fname = InStr(fileList.mybook, Year(date_next) & ".") '�u2020.�v�͉������ڂ���n�܂邩
'    Else                        '1�`11��
'        Fname = InStr(fileList.mybook, Year(date_now) & ".") '�u2020.�v�͉������ڂ���n�܂邩
'    End If

    fileList.this_filename = Mid(fileList.mybook, 1, s - 6) '��   �݌Ɂw�����ޗ��x_

'    If fileList.this_filename = fileList.checkFilename Then
'        fileList.bool_filename = True
'    Else
'        fileList.bool_filename = False
'    End If

    fileList.mybook_month = Int(Mid(str2, 1, s2 - 1)) '8
    '��[   \\Afnewt320-kyoyu\�Г����L\�l�t�H���_\�}��\�݌ɕ\�쐬�}�N��\����\�݌Ɂw�|���x_2020.9.xlsm   ]
    
    fileList.next_mybook = fileList.mypath & "\" & fileList.this_filename & Year(date_next) & "." & Month(date_next) & ".xlsm"

    Set �t�@�C���N���X = fileList
End Function

Public Function �J���[�N���X() As colorClass
    '�錾
    Dim colorList  As colorClass
    Set colorList = New colorClass

    '�z��錾
    Dim a() As Variant
    Dim b() As Variant
    '�������̔z����쐬
    a = Array(6629, 6662, 6672, 6688, 6704, 6714, 6536, 6735, 6684, 6695, 6712, 6673)
    b = Array("���r������", "���{�H����")

    '�Z�b�g
    colorList.colorGenArray = a
    colorList.colorSheetArray = b
    
    Set �J���[�N���X = colorList

End Function
