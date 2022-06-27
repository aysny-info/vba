Sub csv_main()

    '�A�N�e�B�u
    Workbooks("�����a���߂�.xlsm").Activate
    
    Application.ScreenUpdating = False                 '��ʒ�~
    Application.Calculation = xlCalculationManual     '�蓮�v�Z
'    ActiveSheet.Unprotect      '�ی����
    
    '�t�@�C���擾�a��
    csvFilePath_aduke = "\\Afnewt320-kyoyu\�Г����L\AFSKS\�݌ɕ\\csv\�a��\"
    file_list_aduke = csv�t�@�C�����T��(csvFilePath_aduke)
    
    '�t�@�C���擾�߂�
    csvFilePath_modoshi = "\\Afnewt320-kyoyu\�Г����L\AFSKS\�݌ɕ\\csv\�߂�\"
    file_list_modoshi = csv�t�@�C�����T��(csvFilePath_modoshi)
    
    'csv�\��t��_�a��
    sh_name = "�a��csv"
    Dim aduke_data As Variant
    aduke_data = getCSV_utf8(sh_name, file_list_aduke, csvFilePath_aduke)
    
    'csv�\��t��_�߂�
    sh_name = "�߂�csv"
    Dim modoshi_data As Variant
    modoshi_data = getCSV_utf8(sh_name, file_list_modoshi, csvFilePath_modoshi)
    
    Worksheets("�ړ�����").Activate
    
    
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

    date_G2 = Worksheets("�ړ�����").Range("G2") '���t
    Dim paste_one_aduke As Variant  '�\��t���f�[�^
    Dim paste_one_modoshi As Variant  '�\��t���f�[�^

    Workbooks("�����a���߂�.xlsm").Activate
    '*****************�a���f�[�^�擾************************************
    Worksheets("�a��csv").Activate
    
    Dim LastRow As Long '�ŏI�s�擾
    LastRow = Cells(Rows.Count, 2).End(xlUp).Row
    
    Dim LastCol As Long '�ŏI��擾
    LastCol = Cells(1, 1).End(xlToRight).Column + 5

    aduke_data = Range(Cells(1, 1), Cells(LastRow, LastCol))    '�u�a��csv�v�V�[�g�̃f�[�^

    '�i�[����Q�����z��T�C�Y�ݒ�
    ReDim paste_one_aduke(1 To UBound(aduke_data, 1), 1 To UBound(aduke_data, 2)) '(�s,��)
    r = 1
    For i = 1 To UBound(aduke_data)
        If i = 1 Then
            For c = 1 To UBound(aduke_data, 2)
                paste_one_aduke(r, c) = aduke_data(i, c)
            Next c
            r = r + 1
        ElseIf aduke_data(i, 10) = "�a��" And aduke_data(i, 12) = date_G2 Then
            For c = 1 To UBound(aduke_data, 2)
                paste_one_aduke(r, 1) = r - 1      'A��No
                paste_one_aduke(r, c) = aduke_data(i, c)
            Next c
            r = r + 1
        End If
    Next i
    
    '�N���A���ē\��t��
    Worksheets("�a���`��").Activate
    Worksheets("�a���`��").Cells.ClearContents
    Range(Cells(1, 1), Cells(UBound(paste_one_aduke, 1), UBound(paste_one_aduke, 2))) = paste_one_aduke
    
    '*****************�߂��f�[�^�擾************************************
    Worksheets("�߂�csv").Activate
    
    'Dim LastRow As Long '�ŏI�s�擾
    LastRow = Cells(Rows.Count, 2).End(xlUp).Row
    
    'Dim LastCol As Long '�ŏI��擾
    LastCol = Cells(1, 1).End(xlToRight).Column + 5

    aduke_data = Range(Cells(1, 1), Cells(LastRow, LastCol))    '�u�a��csv�v�V�[�g�̃f�[�^

    '�i�[����Q�����z��T�C�Y�ݒ�
    ReDim paste_one_modoshi(1 To UBound(aduke_data, 1), 1 To UBound(aduke_data, 2)) '(�s,��)
    r = 1
    For i = 1 To UBound(aduke_data)
        If i = 1 Then
            For c = 1 To UBound(aduke_data, 2)
                paste_one_modoshi(r, c) = aduke_data(i, c)
            Next c
            r = r + 1
        ElseIf aduke_data(i, 10) = "�߂�" And aduke_data(i, 12) = date_G2 Then
            For c = 1 To UBound(aduke_data, 2)
                paste_one_modoshi(r, 1) = r - 1      'A��No
                paste_one_modoshi(r, c) = aduke_data(i, c)
            Next c
            r = r + 1
        End If
    Next i
    
    '�N���A���ē\��t��
    Worksheets("�߂��`��").Activate
    Worksheets("�߂��`��").Cells.ClearContents
    Range(Cells(1, 1), Cells(UBound(paste_one_modoshi, 1), UBound(paste_one_modoshi, 2))) = paste_one_modoshi


    '******************************* �e�ꏊ *****************************************
    '�ݑq�ɓ\��t��
    sh_name = "�ݑq��"
    Call filter_paste(sh_name, paste_one_aduke, paste_one_modoshi)

    '�X�[�p�[���b�N�X�\��t��
    sh_name = "�X�[�p�[���b�N�X"
    Call filter_paste(sh_name, paste_one_aduke, paste_one_modoshi)

    '�V�؏���
    sh_name = "�V�؏���"
    Call filter_paste(sh_name, paste_one_aduke, paste_one_modoshi)

    '�^�h�R������
    sh_name = "�^�h�R������"
    Call filter_paste(sh_name, paste_one_aduke, paste_one_modoshi)

    '���Ѓg���b�N
    sh_name = "���Ѓg���b�N"
    Call filter_paste(sh_name, paste_one_aduke, paste_one_modoshi)
    
    '�S�V�[�g�Čv�Z
    Application.Calculate
    
    '�w����t�Ńt�B���^�[
    Worksheets("�ړ�����").Activate
    Call �t�B���^�[�N���A
    Range("E4").AutoFilter Field:=3, Criteria1:="<>"
    
    Worksheets("�ݑq��").Activate
    Call �t�B���^�[�N���A
    Range("D6").AutoFilter Field:=11, Criteria1:="<>"

    Worksheets("�X�[�p�[���b�N�X").Activate
    Call �t�B���^�[�N���A
    Range("D6").AutoFilter Field:=11, Criteria1:="<>"

    Worksheets("�V�؏���").Activate
    Call �t�B���^�[�N���A
    Range("D6").AutoFilter Field:=11, Criteria1:="<>"

    Worksheets("�^�h�R������").Activate
    Call �t�B���^�[�N���A
    Range("D6").AutoFilter Field:=11, Criteria1:="<>"

    Worksheets("���Ѓg���b�N").Activate
    Call �t�B���^�[�N���A
    Range("D7").AutoFilter Field:=11, Criteria1:="<>"
    
    '�߂�
    Worksheets("�ړ�����").Activate
    
    Call �p���b�g�폜
    
    Application.ScreenUpdating = True                 '��ʒ�~
    Application.Calculation = xlCalculationAutomatic       '�蓮�v�Z
    ActiveSheet.Protect      '�ی�
    
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
    date_G2 = Worksheets("�ړ�����").Range("G2") '���t
    
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
    
    For n = 1 To UBound(file_list)
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
    
    For n = 1 To UBound(file_list)
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

Function csv�t�@�C�����T��(csvFilePath As Variant) As Variant
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
            filename(cnt, 0) = f.Name
            filename(cnt, 1) = f.DateLastModified   '�쐬����f.DateCreated �X�V����f.DateLastModified  �A�N�Z�X����f.DateLastAccessed
            cnt = cnt + 1
        Next f
    End With
    
    csv�t�@�C�����T�� = filename
    
End Function






