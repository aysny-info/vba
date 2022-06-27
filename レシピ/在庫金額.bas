Sub csv_main()

    '�A�N�e�B�u
    Workbooks("�݌ɋ��z.xlsm").Activate
    
    Application.ScreenUpdating = False                 '��ʒ�~
    Application.Calculation = xlCalculationManual     '�蓮�v�Z
'    ActiveSheet.Unprotect      '�ی����
    
    '�t�@�C���擾�݌ɐ�
    csvFilePath_aduke = "\\Afnewt320-kyoyu\�Г����L\AFSKS\�݌ɕ\\csv\�݌ɐ�\"
    file_list_aduke = csv�t�@�C�����T��(csvFilePath_aduke)
    
    '�t�@�C���擾�O���݌ɐ�
    csvFilePath_modoshi = "\\Afnewt320-kyoyu\�Г����L\AFSKS\�݌ɕ\\csv\�O���݌ɐ�\"
    file_list_modoshi = csv�t�@�C�����T��(csvFilePath_modoshi)
    
    'csv�\��t��_�݌ɐ�
    sh_name = "�݌ɐ�csv"
    Dim zaiko_data As Variant
    zaiko_data = getCSV_utf8(sh_name, file_list_aduke, csvFilePath_aduke)
    
    'csv�\��t��_�O���݌ɐ�
    sh_name = "�O���݌ɐ�csv"
    Dim modoshi_data As Variant
    modoshi_data = getCSV_utf8(sh_name, file_list_modoshi, csvFilePath_modoshi)
    
    'one�X�V
    Call one_search
    
    Worksheets("�I�����ו\").Activate
    
    Application.ScreenUpdating = True                 '��ʒ�~
    Application.Calculation = xlCalculationAutomatic       '�蓮�v�Z
End Sub

Sub tomorrow_add()
    Range("J3").Select
    Range("J3").Value = DateAdd("d", 1, Range("J3"))
    Call one_search
    
End Sub

Sub yesterday_add()
    Range("J3").Select
    Range("J3").Value = DateAdd("d", -1, Range("J3"))
    Call one_search
End Sub

Sub one_search()
    Application.ScreenUpdating = False                 '��ʒ�~
    Application.Calculation = xlCalculationManual     '�蓮�v�Z
    ActiveSheet.Unprotect      '�ی����

    date_G2 = Worksheets("�I�����ו\").Range("J3") '���t
    Dim paste_one_zaiko As Variant  '�\��t���f�[�^
    Dim paste_one_kowake As Variant  '�\��t���f�[�^

    Workbooks("�݌ɋ��z.xlsm").Activate
    '*****************�݌ɐ��f�[�^�擾************************************
    Worksheets("�݌ɐ�csv").Activate
    
    Dim lastRow As Long '�ŏI�s�擾
    lastRow = Cells(Rows.Count, 2).End(xlUp).Row
    
    Dim LastCol As Long '�ŏI��擾
    LastCol = Cells(1, 1).End(xlToRight).Column + 5

    zaiko_data = Range(Cells(1, 1), Cells(lastRow, LastCol))    '�u�݌ɐ�csv�v�V�[�g�̃f�[�^
    
    
    '*****************�O���݌ɐ��f�[�^�擾************************************
    Worksheets("�O���݌ɐ�csv").Activate
    
    
    'Dim LastRow As Long '�ŏI�s�擾
    lastRow = Cells(Rows.Count, 2).End(xlUp).Row
    
    'Dim LastCol As Long '�ŏI��擾
    LastCol = Cells(1, 1).End(xlToRight).Column + 5

    gaibu_row = Range(Cells(1, 1), Cells(lastRow, LastCol))    '�u�O���݌ɐ�csv�v�V�[�g�̃f�[�^

    '�i�[����Q�����z��T�C�Y�ݒ�
    ReDim paste_one_zaiko(1 To (UBound(zaiko_data, 1) + UBound(gaibu_row, 1)), 1 To UBound(zaiko_data, 2))    '(�s,��)
    ReDim paste_one_kowake(1 To (UBound(zaiko_data, 1) + UBound(gaibu_row, 1)), 1 To UBound(zaiko_data, 2))  '(�s,��)

    r = 1
    k = 1
    For i = 1 To UBound(zaiko_data)
        If i = 1 Then
            For c = 1 To UBound(zaiko_data, 2)
                paste_one_zaiko(r, c) = zaiko_data(i, c)
                paste_one_kowake(k, c) = zaiko_data(k, c)
            Next c
            r = r + 1   '�ʏ�݌ɐ�
            k = k + 1   '������
        ElseIf zaiko_data(i, 10) = "�݌ɐ�" And zaiko_data(i, 12) = date_G2 Then
            If zaiko_data(i, 2) Like "*�������i*" Then
                For c = 1 To UBound(zaiko_data, 2)
                    paste_one_kowake(k, 1) = k - 1      'A��No
                    paste_one_kowake(k, c) = zaiko_data(i, c)
                Next c
                k = k + 1
            Else
                For c = 1 To UBound(zaiko_data, 2)
                    paste_one_zaiko(r, 1) = r - 1      'A��No
                    paste_one_zaiko(r, c) = zaiko_data(i, c)
                Next c
                r = r + 1
            End If
        End If
    Next i
    
    '*****************�O���݌ɐ��f�[�^�擾************************************
    Worksheets("�O���݌ɐ�csv").Activate
    
    'Dim LastRow As Long '�ŏI�s�擾
    lastRow = Cells(Rows.Count, 2).End(xlUp).Row
    
    'Dim LastCol As Long '�ŏI��擾
    LastCol = Cells(1, 1).End(xlToRight).Column + 5

    zaiko_data = Range(Cells(1, 1), Cells(lastRow, LastCol))    '�u�O���݌ɐ�csv�v�V�[�g�̃f�[�^

'    r = 1
'    k = 1
    
    For i = 1 To UBound(zaiko_data, 1)
        If i = 1 Then
            ' For c = 1 To UBound(zaiko_data, 2)
            '     paste_one_zaiko(r, c) = zaiko_data(i, c)
            '     paste_one_kowake(k, c) = zaiko_data(k, c)
            ' Next c
'            r = r + 1   '�ʏ�݌ɐ�
'            k = k + 1   '������
        ElseIf zaiko_data(i, 10) = "�O���݌ɐ�" And zaiko_data(i, 12) = date_G2 Then
            If zaiko_data(i, 2) Like "*�������i*" Then
                For c = 1 To UBound(zaiko_data, 2)
                    paste_one_kowake(k, 1) = k - 1      'A��No
                    paste_one_kowake(k, c) = zaiko_data(i, c)
                Next c
                k = k + 1
            Else
                For c = 1 To UBound(zaiko_data, 2)
                    paste_one_zaiko(r, 1) = r - 1      'A��No
                    paste_one_zaiko(r, c) = zaiko_data(i, c)
                Next c
                r = r + 1
            End If
        End If
    Next i
    
    '�N���A���ē\��t��
    Worksheets("�݌ɐ��`��").Activate
    Worksheets("�݌ɐ��`��").Cells.ClearContents
    Range(Cells(1, 1), Cells(UBound(paste_one_zaiko, 1), UBound(paste_one_zaiko, 2))) = paste_one_zaiko
     '�N���A���ē\��t��
    Worksheets("�������݌ɐ��`��").Activate
    Worksheets("�������݌ɐ��`��").Cells.ClearContents
    Range(Cells(1, 1), Cells(UBound(paste_one_kowake, 1), UBound(paste_one_kowake, 2))) = paste_one_kowake
 
    '�S�V�[�g�Čv�Z
    Application.Calculate
     
    '�߂�
    Worksheets("�I�����ו\").Activate
    
    Application.ScreenUpdating = True                 '��ʋN��
    Application.Calculation = xlCalculationAutomatic       '�����v�Z
    ActiveSheet.Protect AllowFiltering:=True       '�ی�
    
End Sub

Function getCSV_utf8(sh_name As Variant, file_list As Variant, csvFilePath As Variant) As Variant
    
    'Dim ws As Worksheet
    'Set ws = ThisWorkbook.Worksheets(1)
    
    'Dim strPath As String
    'strPath = "\\Afnewt320-kyoyu\�Г����L\AFSKS\�݌ɕ\\csv\�݌ɐ�\�y�݌ɐ��z�݌�_�_���{�[��_2022.3.xlsm.csv"
    
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

Function csv�t�@�C�����T��(csvFilePath As Variant) As Variant
    'csvFilePath = "\\Afnewt320-kyoyu\�Г����L\AFSKS\�݌ɕ\\csv\�݌ɐ�"
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


