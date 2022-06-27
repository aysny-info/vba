Sub csv_main()

    '�A�N�e�B�u
    ThisWorkbook.Activate
    
    Application.ScreenUpdating = False                 '��ʒ�~
    Application.Calculation = xlCalculationManual     '�蓮�v�Z
    'ActiveSheet.Unprotect      '�ی����
    
    ' ************************    �R�[�v  ************************************
    '�t�@�C���擾 �v�搔
    csvFilePath_cope = "\\Afnewt320-kyoyu\�Г����L\�y���Y�Ǘ��z\�y�V�X�e���z\csv\�h\"
    file_list_cope = csv�t�@�C�����T��(csvFilePath_cope)
    
    '�R�[�v
    sh_name = "�R�[�v�v�搔"
    Dim cope_data As Variant
    cope_data = getCSV_utf8(sh_name, file_list_cope, csvFilePath_cope)
    
    ' ************************    ���[�R�[�v  ************************************
    '�t�@�C���擾 �v�搔
    csvFilePath_Ucope = "\\Afnewt320-kyoyu\�Г����L\�y���Y�Ǘ��z\�y�V�X�e���z\csv\�p\"
    file_list_Ucope = csv�t�@�C�����T��(csvFilePath_Ucope)
    
    '���[�R�[�v
    sh_name = "���[�R�[�v�v�搔"
    Dim Ucope_data As Variant
    Ucope_data = getCSV_utf8(sh_name, file_list_Ucope, csvFilePath_Ucope)
    
    ' ************************    �߂�  ************************************
    '�t�@�C���擾 �v�搔
    csvFilePath_modoshi = "\\Afnewt320-kyoyu\�Г����L\AFSKS\�݌ɕ\\csv\�߂�\"
    file_list_modoshi = csv�t�@�C�����T��(csvFilePath_modoshi, "���")
    
    '�߂�
    sh_name = "�߂�"
    Dim modoshi_data As Variant
    modoshi_data = getCSV_utf8(sh_name, file_list_modoshi, csvFilePath_modoshi)

    ' ************************    ���א�  ************************************
    '�t�@�C���擾 �v�搔
    csvFilePath_nyuuka = "\\Afnewt320-kyoyu\�Г����L\AFSKS\�݌ɕ\\csv\���א�\"
    file_list_nyuuka = csv�t�@�C�����T��(csvFilePath_nyuuka, "���")
    
    '���א�
    sh_name = "���א�"
    Dim nyuuka_data As Variant
    nyuuka_data = getCSV_utf8(sh_name, file_list_nyuuka, csvFilePath_nyuuka)

    ' ************************    �݌ɐ�  ************************************
    '�t�@�C���擾 �݌ɐ�
    csvFilePath_zaiko = "\\Afnewt320-kyoyu\�Г����L\AFSKS\�݌ɕ\\csv\�݌ɐ�\"
    file_list_zaiko = csv�t�@�C�����T��(csvFilePath_zaiko, "���")
    
    '�݌ɐ�
    sh_name = "�݌ɐ�"
    Dim zaiko_data As Variant
    zaiko_data = getCSV_utf8(sh_name, file_list_zaiko, csvFilePath_zaiko)

    Worksheets("���V�s�\��\").Activate
    
    Application.ScreenUpdating = True                 '��ʒ�~
    Application.Calculation = xlCalculationAutomatic       '�蓮�v�Z
End Sub

Sub tomorrow_add()
    Range("G2").Select
    Range("G2").Value = DateAdd("d", 1, Range("G2"))
    Call one_search
    
End Sub

Sub yesterday_add()
    Range("G2").Select
    Range("G2").Value = DateAdd("d", -1, Range("G2"))
    Call one_search
End Sub

Sub one_search()
    Application.ScreenUpdating = False                 '��ʒ�~
    Application.Calculation = xlCalculationManual     '�蓮�v�Z
    ActiveSheet.Unprotect      '�ی����

    ThisWorkbook.Activate

    date_G2 = Worksheets("���V�s�\��\").Range("M18") '���t
    Dim paste_one_aduke As Variant  '�\��t���f�[�^
    Dim paste_one_modoshi As Variant  '�\��t���f�[�^
    Dim LastRow As Long '�ŏI�s�擾
    Dim LastCol As Long '�ŏI��擾
    
    '*****************�R�[�v�擾************************************
    Worksheets("�R�[�v�v�搔").Activate
    
    LastRow = Cells(Rows.Count, 2).End(xlUp).Row
    
    LastCol = Cells(1, 1).End(xlToRight).Column + 5

    cope_data = Range(Cells(1, 1), Cells(LastRow, LastCol))    '�u�R�[�v�v�搔�v�V�[�g�̃f�[�^

    '�i�[����Q�����z��T�C�Y�ݒ�
    ReDim paste_one_cope(1 To UBound(cope_data, 1), 1 To UBound(cope_data, 2) + 2) '(�s,��)
    r = 1

    CP_hiduke = DateAdd("d", 5, date_G2) '���j������T����͓y�j���i�����j
    column_b = True
    No_cope = 1

    For d = 1 To 5
        For i = 1 To UBound(cope_data)
            If i = 1 And column_b Then
                For c = 1 To UBound(cope_data, 2)
                    paste_one_cope(r, 1) = "No"
                    paste_one_cope(r, c + 1) = cope_data(i, c)
                Next c
                r = r + 1
                column_b = False
            ElseIf cope_data(i, 2) = CP_hiduke Then
                For c = 1 To UBound(cope_data, 2)
                    'paste_one_cope(r, 1) = Trim(str(d) & Right("00" & str(r - 1), 2)) 'A��No
                    paste_one_cope(r, 1) = No_cope 'A��No
                    paste_one_cope(r, c + 1) = cope_data(i, c)
                Next c
                r = r + 1
                No_cope = No_cope + 1
            End If
        Next i
        CP_hiduke = DateAdd("d", 1, CP_hiduke)
        No_cope = 1
    Next d

    '�N���A���ē\��t��
    Worksheets("�R�[�v�`��").Activate
    ActiveSheet.Unprotect      '�ی����
    Worksheets("�R�[�v�`��").Cells.ClearContents
    Range(Cells(1, 1), Cells(UBound(paste_one_cope, 1), UBound(paste_one_cope, 2))) = paste_one_cope
    
    '*****************���[�R�[�v�擾************************************
    Worksheets("���[�R�[�v�v�搔").Activate
    
    LastRow = Cells(Rows.Count, 2).End(xlUp).Row
    
    LastCol = Cells(1, 1).End(xlToRight).Column + 5

    Ucope_data = Range(Cells(1, 1), Cells(LastRow, LastCol))    '�u���[�R�[�v�v�搔�v�V�[�g�̃f�[�^

    '�i�[����Q�����z��T�C�Y�ݒ�
    ReDim paste_one_udata(1 To UBound(Ucope_data, 1), 1 To UBound(Ucope_data, 2) + 2) '(�s,��)
    r = 1

    CP_hiduke = DateAdd("d", 5, date_G2) '���j������T����͓y�j���i�����j
    column_b = True
    No_cope = 1

    For d = 1 To 5
        For i = 1 To UBound(Ucope_data)
            If i = 1 And column_b Then
                For c = 1 To UBound(Ucope_data, 2)
                    paste_one_udata(r, 1) = "No"
                    paste_one_udata(r, c + 1) = Ucope_data(i, c)
                Next c
                r = r + 1
                column_b = False
            ElseIf Ucope_data(i, 2) = CP_hiduke Then
                For c = 1 To UBound(Ucope_data, 2)
                    paste_one_udata(r, 1) = No_cope     'A��No
                    paste_one_udata(r, c + 1) = Ucope_data(i, c)
                Next c
                r = r + 1
                No_cope = No_cope + 1
            End If
        Next i
        CP_hiduke = DateAdd("d", 1, CP_hiduke)
        No_cope = 1
    Next d
    '�N���A���ē\��t��
    Worksheets("���[�R�[�v�`��").Activate
    ActiveSheet.Unprotect      '�ی����
    Worksheets("���[�R�[�v�`��").Cells.ClearContents
    Range(Cells(1, 1), Cells(UBound(paste_one_udata, 1), UBound(paste_one_udata, 2))) = paste_one_udata

    '*****************�߂��擾************************************
    Worksheets("�߂�").Activate
    
    LastRow = Cells(Rows.Count, 2).End(xlUp).Row
    
    LastCol = Cells(1, 1).End(xlToRight).Column + 5

    modoshi_data = Range(Cells(1, 1), Cells(LastRow, LastCol))    '�u�߂��v�V�[�g�̃f�[�^

    '�i�[����Q�����z��T�C�Y�ݒ�
    ReDim paste_one_modoshi(1 To UBound(modoshi_data, 1), 1 To UBound(modoshi_data, 2) + 2) '(�s,��)
    r = 1

    CP_hiduke = date_G2 '���j������T����͓y�j���i�����j
    column_b = True
    No_cope = 1

    For d = 1 To 10
        For i = 1 To UBound(modoshi_data)
            If i = 1 And column_b Then
                For c = 1 To UBound(modoshi_data, 2)
                    paste_one_modoshi(r, 1) = "No"
                    paste_one_modoshi(r, c + 1) = modoshi_data(i, c)
                Next c
                r = r + 1
                column_b = False
            ElseIf modoshi_data(i, 12) = CP_hiduke Then
                For c = 1 To UBound(modoshi_data, 2)
                    paste_one_modoshi(r, 1) = No_cope    'A��No
                    paste_one_modoshi(r, c + 1) = modoshi_data(i, c)
                Next c
                r = r + 1
                No_cope = No_cope + 1
            End If
        Next i
        CP_hiduke = DateAdd("d", 1, CP_hiduke)
        No_cope = 1
    Next d
    
    '�N���A���ē\��t��
    Worksheets("�߂��`��").Activate
    ActiveSheet.Unprotect      '�ی����
    Worksheets("�߂��`��").Cells.ClearContents
    Range(Cells(1, 1), Cells(UBound(paste_one_modoshi, 1), UBound(paste_one_modoshi, 2))) = paste_one_modoshi

    '*****************���א��擾************************************
    Worksheets("���א�").Activate
    
    LastRow = Cells(Rows.Count, 2).End(xlUp).Row
    
    LastCol = Cells(1, 1).End(xlToRight).Column + 5

    nyuuka_data = Range(Cells(1, 1), Cells(LastRow, LastCol))    '�u���א��v�V�[�g�̃f�[�^

    '�i�[����Q�����z��T�C�Y�ݒ�
    ReDim paste_one_nyuuka(1 To UBound(nyuuka_data, 1), 1 To UBound(nyuuka_data, 2) + 2) '(�s,��)
    r = 1

    CP_hiduke = date_G2 '���j������T����͓y�j���i�����j
    column_b = True
    No_cope = 1

    For d = 1 To 10
        For i = 1 To UBound(nyuuka_data)
            If i = 1 And column_b Then
                For c = 1 To UBound(nyuuka_data, 2)
                    paste_one_nyuuka(r, 1) = "No"
                    paste_one_nyuuka(r, c + 1) = nyuuka_data(i, c)
                Next c
                r = r + 1
                column_b = False
            ElseIf nyuuka_data(i, 12) = CP_hiduke Then
                For c = 1 To UBound(nyuuka_data, 2)
                    paste_one_nyuuka(r, 1) = No_cope    'A��No
                    paste_one_nyuuka(r, c + 1) = nyuuka_data(i, c)
                Next c
                r = r + 1
                No_cope = No_cope + 1
            End If
        Next i
        CP_hiduke = DateAdd("d", 1, CP_hiduke)
        No_cope = 1
    Next d
    
    '�N���A���ē\��t��
    Worksheets("���א��`��").Activate
    ActiveSheet.Unprotect      '�ی����
    Worksheets("���א��`��").Cells.ClearContents
    Range(Cells(1, 1), Cells(UBound(paste_one_nyuuka, 1), UBound(paste_one_nyuuka, 2))) = paste_one_nyuuka


    '*****************�݌ɐ��擾************************************
    Worksheets("�݌ɐ�").Activate
    
    LastRow = Cells(Rows.Count, 2).End(xlUp).Row
    
    LastCol = Cells(1, 1).End(xlToRight).Column + 5

    zaiko_data = Range(Cells(1, 1), Cells(LastRow, LastCol))    '�u�݌ɐ��v�V�[�g�̃f�[�^

    '�i�[����Q�����z��T�C�Y�ݒ�
    ReDim paste_one_zaiko(1 To UBound(zaiko_data, 1), 1 To UBound(zaiko_data, 2) + 2) '(�s,��)
    r = 1

    CP_hiduke = date_G2 '���j������T����͓y�j���i�����j
    column_b = True
    No_cope = 1

    For d = 1 To 10
        For i = 1 To UBound(zaiko_data)
            If i = 1 And column_b Then
                For c = 1 To UBound(zaiko_data, 2)
                    paste_one_zaiko(r, 1) = "No"
                    paste_one_zaiko(r, c + 1) = zaiko_data(i, c)
                Next c
                r = r + 1
                column_b = False
            ElseIf zaiko_data(i, 12) = CP_hiduke Then
                For c = 1 To UBound(zaiko_data, 2)
                    paste_one_zaiko(r, 1) = No_cope    'A��No
                    paste_one_zaiko(r, c + 1) = zaiko_data(i, c)
                Next c
                r = r + 1
                No_cope = No_cope + 1
            End If
        Next i
        CP_hiduke = DateAdd("d", 1, CP_hiduke)
        No_cope = 1
    Next d
    
    '�N���A���ē\��t��
    Worksheets("�݌ɐ��`��").Activate
    ActiveSheet.Unprotect      '�ی����
    Worksheets("�݌ɐ��`��").Cells.ClearContents
    Range(Cells(1, 1), Cells(UBound(paste_one_zaiko, 1), UBound(paste_one_zaiko, 2))) = paste_one_zaiko

    Application.ScreenUpdating = True                 '��ʒ�~
    Application.Calculation = xlCalculationAutomatic       '�蓮�v�Z
    'ActiveSheet.Protect      '�ی�
    Worksheets("���V�s�\��\").Activate
    
End Sub
Sub paste_one_data()
    Application.ScreenUpdating = False                 '��ʒ�~
    Application.Calculation = xlCalculationManual     '�蓮�v�Z

    ThisWorkbook.Activate
    Worksheets("���V�s�\��\").Activate
    Worksheets("���V�s�\��\").Unprotect
    
    Application.Calculate '�Čv�Z
    
    ' ************************    �f�[�^�ϐ��֊i�[  ************************************
    '���V�sNo
    u_recipe_no = Range(Cells(1, 26), Cells(15, 26))
    c_recipe_no = Range(Cells(18, 26), Cells(41, 26))
    
    '�v�搔
    u_syo = Range(Cells(1, 7), Cells(15, 7))
    c_syo = Range(Cells(18, 7), Cells(41, 7))

    '�߂��y�[�X�g�f�[�^
    Range(Cells(2, 30), Cells(15, 39)).ClearContents
    Range(Cells(19, 30), Cells(41, 39)).ClearContents
    modoshi_U_paste = Range(Cells(1, 30), Cells(15, 39)).Formula
    modoshi_U_date = Range(Cells(1, 30), Cells(1, 39))              '��̓��t
    modoshi_C_paste = Range(Cells(18, 30), Cells(41, 39)).Formula
    modoshi_C_date = Range(Cells(18, 30), Cells(18, 39))            '��̓��t

    '���׃y�[�X�g�f�[�^
    Range(Cells(2, 41), Cells(15, 50)).ClearContents
    Range(Cells(19, 41), Cells(41, 50)).ClearContents
    nyuuka_U_paste = Range(Cells(1, 41), Cells(15, 50)).Formula
    nyuuka_U_date = Range(Cells(1, 41), Cells(1, 50))              '��̓��t
    nyuuka_C_paste = Range(Cells(18, 41), Cells(41, 50)).Formula
    nyuuka_C_date = Range(Cells(18, 41), Cells(18, 50))            '��̓��t
    '�v�搔�y�[�X�g�f�[�^
    Range(Cells(2, 52), Cells(15, 56)).ClearContents
    Range(Cells(19, 52), Cells(41, 56)).ClearContents
    keikaku_U_paste = Range(Cells(1, 52), Cells(15, 56)).Formula
    keikaku_U_date = Range(Cells(1, 52), Cells(1, 56))              '��̓��t
    keikaku_C_paste = Range(Cells(18, 52), Cells(41, 56)).Formula
    keikaku_C_date = Range(Cells(18, 52), Cells(18, 56))            '��̓��t
    
    ' ************************    �y�[�X�g�f�[�^���W  ************************************
    '�߂��y�[�X�g�f�[�^
    modoshi_U_paste = get_paste_data(u_recipe_no, modoshi_U_paste, "�߂��`��", modoshi_U_date)
    modoshi_C_paste = get_paste_data(c_recipe_no, modoshi_C_paste, "�߂��`��", modoshi_C_date)
    '���׃y�[�X�g�f�[�^
    nyuuka_U_paste = get_paste_data(u_recipe_no, nyuuka_U_paste, "���א��`��", nyuuka_U_date)
    nyuuka_C_paste = get_paste_data(c_recipe_no, nyuuka_C_paste, "���א��`��", nyuuka_C_date)
    '�v�搔�y�[�X�g�f�[�^
    keikaku_U_paste = get_keikaku(u_syo, keikaku_U_paste, "���[�R�[�v�`��", keikaku_U_date)
    keikaku_C_paste = get_keikaku(c_syo, keikaku_C_paste, "�R�[�v�`��", keikaku_C_date)

    ' ************************    �y�[�X�g  ************************************
    Worksheets("���V�s�\��\").Activate
    '�߂�
    Range(Cells(1, 30), Cells(15, 39)) = modoshi_U_paste
    Range(Cells(18, 30), Cells(41, 39)) = modoshi_C_paste
    '����
    Range(Cells(1, 41), Cells(15, 50)) = nyuuka_U_paste
    Range(Cells(18, 41), Cells(41, 50)) = nyuuka_C_paste
    
    '�v��
     Range(Cells(1, 52), Cells(15, 56)) = keikaku_U_paste
     Range(Cells(18, 52), Cells(41, 56)) = keikaku_C_paste

    Application.ScreenUpdating = True                 '��ʒ�~
    Application.Calculation = xlCalculationAutomatic       '�蓮�v�Z

End Sub

Function get_paste_data(recipe_no As Variant, paste_data As Variant, sheet_name As Variant, ueno_date) As Variant
    Worksheets(sheet_name).Activate
    Dim this_sh_data As Variant  '�\��t���f�[�^
    Dim LastRow As Long '�ŏI�s�擾
    Dim LastCol As Long '�ŏI��擾
    LastRow = Cells(Rows.Count, 3).End(xlUp).Row
    LastCol = Cells(1, 1).End(xlToRight).Column + 5
    
    this_sh_data = Range(Cells(1, 1), Cells(LastRow, LastCol))
    
    For r = 1 To UBound(recipe_no)  '���V�sNo��
        For t = 1 To UBound(this_sh_data)   '�`���f�[�^����
            If this_sh_data(t, 7) = recipe_no(r, 1) Then    '���V�sNo�ƌ`���f�[�^��7���
                For p = 1 To UBound(paste_data, 2)  '�y�[�X�g�̗񐔉�
                    If ueno_date(1, p) = this_sh_data(t, 13) Then
                        If sheet_name Like "*�߂��`��*" Then
                            paste_data(r, p) = "�ݑq�� " & this_sh_data(t, 12)
                        ElseIf sheet_name Like "*���א��`��*" Then
                            paste_data(r, p) = "���א� " & this_sh_data(t, 12)
                        End If
                    End If
                Next p
            End If
        Next t
    Next r
    
    get_paste_data = paste_data

End Function

Function get_keikaku(keikaku_no As Variant, paste_data As Variant, sheet_name As Variant, ueno_date) As Variant
    Worksheets(sheet_name).Activate
    Dim this_sh_data As Variant  '�\��t���f�[�^
    Dim LastRow As Long '�ŏI�s�擾
    Dim LastCol As Long '�ŏI��擾
    LastRow = Cells(Rows.Count, 3).End(xlUp).Row
    LastCol = Cells(1, 1).End(xlToRight).Column + 5
    
    this_sh_data = Range(Cells(1, 1), Cells(LastRow, LastCol))
    
    For r = 1 To UBound(keikaku_no)  '���i�R�[�h��
        For t = 1 To UBound(this_sh_data)   '�`���f�[�^����
            If this_sh_data(t, 2) = keikaku_no(r, 1) Then    '���i�R�[�h�ƌ`���f�[�^��7���
                For p = 1 To UBound(paste_data, 2)  '�y�[�X�g�̗񐔉�
                    If ueno_date(1, p) = this_sh_data(t, 3) Then
                        paste_data(r, p) = this_sh_data(t, 4)
                    End If
                Next p
            End If
        Next t
    Next r
    
    get_keikaku = paste_data

End Function


Sub test()
    Application.ScreenUpdating = True                 '��ʒ�~
    Application.Calculation = xlCalculationAutomatic       '�蓮�v�Z
End Sub

Sub �p���b�g�폜()
    Worksheets("�ړ�����").Activate
    Range("Q5").Select
    Union(Selection, Range("Q5:Q64")).Select
    Union(Selection, Range("W65:W124")).Select
    Selection.ClearContents
    Range("G2").Select
End Sub

Sub filter_paste(sh_name As Variant, paste_one_aduke As Variant, paste_one_modoshi As Variant)
    Dim paste_d As Variant
    date_G2 = Worksheets("���V�s�\��\").Range("M18") '���t
    
    '*****************�y�a���zsh_name�Ńt�B���^�[************************************
    '�i�[����Q�����z��T�C�Y�ݒ�
    ReDim paste_d(1 To UBound(paste_one_aduke, 1), 1 To UBound(paste_one_aduke, 2)) '(�s,��)

    r = 1
    For i = 1 To UBound(paste_one_aduke)
        If i = 1 Then
            For c = 1 To UBound(paste_one_aduke, 2)
                paste_d(r, c) = paste_one_aduke(i, c)
            Next c
            r = r + 1
        ElseIf paste_one_aduke(i, 13) = sh_name And paste_one_aduke(i, 12) = date_G2 Then
            For c = 1 To UBound(paste_one_aduke, 2)
                paste_d(r, 1) = r - 1      'A��No
                paste_d(r, c) = paste_one_aduke(i, c)
            Next c
            r = r + 1
        End If
    Next i
    
    '�N���A���ē\��t��
    Worksheets(sh_name & "�a��" & "�`��").Activate
    Worksheets(sh_name & "�a��" & "�`��").Cells.ClearContents
    Range(Cells(1, 1), Cells(UBound(paste_d, 1), UBound(paste_d, 2))) = paste_d

    '*****************�y�߂��zsh_name�Ńt�B���^�[************************************
    '�i�[����Q�����z��T�C�Y�ݒ�
    ReDim paste_d(1 To UBound(paste_one_modoshi, 1), 1 To UBound(paste_one_modoshi, 2)) '(�s,��)

    r = 1
    For i = 1 To UBound(paste_one_modoshi)
        If i = 1 Then
            For c = 1 To UBound(paste_one_modoshi, 2)
                paste_d(r, c) = paste_one_modoshi(i, c)
            Next c
            r = r + 1
        ElseIf paste_one_modoshi(i, 13) = sh_name And paste_one_modoshi(i, 12) = date_G2 Then
            For c = 1 To UBound(paste_one_modoshi, 2)
                paste_d(r, 1) = r - 1      'A��No
                paste_d(r, c) = paste_one_modoshi(i, c)
            Next c
            r = r + 1
        End If
    Next i
    
    '�N���A���ē\��t��
    Worksheets(sh_name & "�߂�" & "�`��").Activate
    Worksheets(sh_name & "�߂�" & "�`��").Cells.ClearContents
    Range(Cells(1, 1), Cells(UBound(paste_d, 1), UBound(paste_d, 2))) = paste_d

End Sub

Function getCSV_utf8(sh_name As Variant, file_list As Variant, csvFilePath As Variant) As Variant
    
    'Dim ws As Worksheet
    'Set ws = ThisWorkbook.Worksheets(1)
    
    'Dim strPath As String
    'strPath = "\\Afnewt320-kyoyu\�Г����L\AFSKS\�݌ɕ\\csv\�a��\�y�a���z�݌�_�_���{�[��_2022.3.xlsm.csv"
    
    Dim i As Long, j As Long
    Dim strLine As String
    Dim arrLine As Variant '�J���}��split���Ċi�[
    
    'D��ϐ��錾
    Dim paste_data() As Variant
    
    'ADODB.Stream�I�u�W�F�N�g�𐶐�
    Dim adoSt As Object
    Set adoSt = CreateObject("ADODB.Stream")
    
    '�V�[�g�N���A
    ThisWorkbook.Activate
    Worksheets(sh_name).Activate
    Worksheets(sh_name).Cells.ClearContents
    max_n = 0
    i = 1
    
    For n = 0 To UBound(file_list)
        With adoSt
            .Charset = "UTF-8"        'Stream�ň��������R�[�g��utf-8�ɐݒ�
            .Open                             'Stream���I�[�v��
            .LoadFromFile (csvFilePath & file_list(n, 0)) '�t�@�C������Stream�Ƀf�[�^��ǂݍ���
            
            Do Until .EOS           'Stream�̖����܂ŌJ��Ԃ�
                strLine = .ReadText(adReadLine) 'Stream����1�s��荞��
                arrLine = Split(Replace(replaceColon(strLine), """", ""), ":") 'strLine���J���}�ŋ�؂�arrLine�Ɋi�[

                max_n = max_n + 1
            Loop

            .Close
        End With
    Next n

    '�i�[����Q�����z��T�C�Y�ݒ�
    ReDim paste_data(1 To max_n, 1 To 30) '(�s,��)
    
    csv_column_name = 1 '�J���������P�s�ڂɒǉ�
    
    For n = 0 To UBound(file_list)
        csv_row_num = 1
        
        With adoSt
            .Charset = "UTF-8"        'Stream�ň��������R�[�g��utf-8�ɐݒ�
            .Open                             'Stream���I�[�v��
            .LoadFromFile (csvFilePath & file_list(n, 0)) '�t�@�C������Stream�Ƀf�[�^��ǂݍ���
            
            Do Until .EOS           'Stream�̖����܂ŌJ��Ԃ�
                
                    strLine = .ReadText(adReadLine) 'Stream����1�s��荞��
                    arrLine = Split(Replace(replaceColon(strLine), """", ""), ":") 'strLine���J���}�ŋ�؂�arrLine�Ɋi�[
                    
                    If csv_column_name = 1 Then '�J���������P�s�ڂɒǉ�
                        For j = 0 To UBound(arrLine)
                            paste_data(i, j + 1) = arrLine(j)
                        Next j
                        i = i + 1
                        csv_column_name = 2
                    ElseIf csv_row_num <> 1 Then '�f�[�^�̕�����ǉ�
                        For j = 0 To UBound(arrLine)
                            paste_data(i, j + 1) = arrLine(j)
                        Next j
                        i = i + 1
                    End If
                    
                csv_row_num = csv_row_num + 1
            Loop
        
            .Close
        End With
        
    Next n

    Range(Cells(1, 1), Cells(max_n, 30)) = paste_data

    getCSV_utf8 = paste_data

End Function

'�󂯎����������̃J���}���R�����ɒu��������
'�_�u���N�H�[�e�[�V�����ň͂܂�Ă���J���}�͒u�������Ȃ�
Function replaceColon(ByVal str As String) As String
    
    Dim strTemp As String
    Dim quotCount As Long
    
    Dim l As Long
    For l = 1 To Len(str)  'str�̒��������J��Ԃ�
    
        strTemp = Mid(str, l, 1) 'str���猻�݂�1������؂�o��
    
        If strTemp = """" Then   'strTemp���_�u���N�H�[�e�[�V�����Ȃ�
    
            quotCount = quotCount + 1   '�_�u���N�H�[�e�[�V�����̃J�E���g��1���₷
    
        ElseIf strTemp = "," Then   'strTemp���J���}�Ȃ�
    
            If quotCount Mod 2 = 0 Then   'quotCount��2�̔{���Ȃ�
    
                str = Left(str, l - 1) & ":" & Right(str, Len(str) - l)   '���݂�1�������R�����ɒu��������
    
            End If
    
        End If
    
    Next l
    
    replaceColon = str

End Function

Function csv�t�@�C�����T��(csvFilePath As Variant, Optional csvname_ichibu As Variant = "����") As Variant
    'csvFilePath = "\\Afnewt320-kyoyu\�Г����L\AFSKS\�݌ɕ\\csv\�a��"
    '�f�B���N�g�����݃`�F�b�N
    If Dir(csvFilePath, vbDirectory) = "" Then
        MsgBox "csv�f�B���N�g�������݂��܂���B"
        End
    Else
        Debug.Print "�f�B���N�g�������݂��܂��B"
    End If
    
    Dim f As Object, cnt As Long
    Dim filename() As Variant    '�y0�z�t�@�C����.csv�y1�z�쐬���y2�z�N(�t�@�C����)�y3�z��(�t�@�C����)

    cnt = 0
    s = 0
    '�Q�����z��̊i�[���鐔�����߂�
    With CreateObject("Scripting.FileSystemObject")
        For Each f In .GetFolder(csvFilePath).Files
            s = s + 1
        Next f
    End With
    
    'csv���݃`�F�b�N
    If s = 0 Then
        MsgBox "csv�t�@�C������ł��B"
        End
    End If
    
    ReDim filename(s - 1, 4) '�y0�z�t�@�C����.csv�y1�z�쐬���y2�z�N(�t�@�C����)�y3�z��0����(�t�@�C����)�y4�z��0����(�t�@�C����)
    
    With CreateObject("Scripting.FileSystemObject")
        For Each f In .GetFolder(csvFilePath).Files
            If csvname_ichibu = "����" Then
                filename(cnt, 0) = f.Name
                filename(cnt, 1) = f.DateLastModified   '�쐬����f.DateCreated �X�V����f.DateLastModified  �A�N�Z�X����f.DateLastAccessed
                cnt = cnt + 1
            Else
                If f.Name Like "*" & csvname_ichibu & "*" Then
                    filename(cnt, 0) = f.Name
                    filename(cnt, 1) = f.DateLastModified   '�쐬����f.DateCreated �X�V����f.DateLastModified  �A�N�Z�X����f.DateLastAccessed
                    cnt = cnt + 1
                End If
            End If
        Next f
    End With
    
    ReDim tmp(cnt - 1, 4)
    For i = 0 To cnt - 1
        For x = 0 To LBound(filename)
            tmp(i, x) = filename(i, x)
        Next x
    Next i
    
    csv�t�@�C�����T�� = tmp
    
End Function




