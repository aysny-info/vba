
Function ���Y�\�X�Vmain(Customer_name As Variant)
    
    '�A�N�e�B�u
    Worksheets("�󒍓���").Activate
    Worksheets("�󒍓���").Select
    Range("A6").Select  'A1�͍�Ǝ҂ɂ��L�[�{�[�h�����ŏo�ד����ύX�����g���u�������邽�ߋ֎~ 2021.03.30
    
    Application.ScreenUpdating = False                 '��ʒ�~
    Application.Calculation = xlCalculationManual     '�蓮�v�Z
    ActiveSheet.UnProtect      '�ی����
    
    
    csv_data = csv�ǂݍ���(Customer_name)       '�Y����csv�f�[�^
    ship_date = Sheets("�󒍓���").Range("A1")  '�o�ד�
    
    'A��ŏI�s�擾
    Dim ALastRow As Long
    ALastRow = Cells(Rows.Count, 1).End(xlUp).Row
    
    'A��i�[
    A_column = Range(Cells(1, 1), Cells(ALastRow, 1))

    A_start_row = 0
    A_END_row = 0
    
    'A�񂩂�����������
    For i = 1 To UBound(A_column)
            If A_column(i, 1) = Customer_name Then
                A_start_row = i + 1
                e = i
                Exit For
            End If
    Next i
    'A�񂩂�END������ �����̌�����̍s����J�n
    For i = e To UBound(A_column)
            If A_column(i, 1) = "END" Then
                A_END_row = i - 1
                Exit For
            End If
    Next

    'D��ϐ��錾
    Dim D_column() As Variant
    ReDim D_column(1 To A_END_row - A_start_row + 1, 1 To 1)
    '0�Ŗ��߂�
    For i = 1 To UBound(D_column)
        D_column(i, 1) = 0
    Next i
    
    'D_column = Range(Cells(A_start_row, 4), Cells(A_END_row, 4))    '�錾���ʓ|�Ȃ̂ŃR�s�[���Ď����Ă���
    A_col_code = Range(Cells(A_start_row, 1), Cells(A_END_row, 1))  'A��̏��i�R�[�h�@�����p

    ship_sum = 0    '���v��
    'csv�f�[�^�̏o�א���D��֊i�[
    For i = 0 To UBound(csv_data)
        For a = 1 To UBound(A_col_code)
            If A_col_code(a, 1) = csv_data(i, 0) Then
                D_column(a, 1) = csv_data(i, 1)
                ship_sum = ship_sum + csv_data(i, 1)
            End If
        Next a
    Next i
    
    Debug.Print 2234
    'D��֓\�t
    Range(Cells(A_start_row, 4), Cells(A_END_row, 4)) = D_column
    
    MsgBox "�X�V�������܂����B" & vbCrLf & vbCrLf & "�y�o�א�z" & Customer_name & vbCrLf & "�y�o�ד��z" & ship_date & vbCrLf & "�y���v���z" & ship_sum & vbCrLf & vbCrLf & "��OK�{�^���N���b�N��V�[�g�̍Čv�Z���s���܂��B" & vbCrLf & "�E�����P�O�O���ɂȂ�܂ł��҂��������B"
    
    If Customer_name = "CGC" Then
        'pass
    ElseIf 1 = 1 Then
        Application.ScreenUpdating = True                  '��ʋN��
        Application.Calculation = xlCalculationAutomatic  '�����v�Z
        ActiveSheet.Protect       '�ی�
    End If
        
'    Debug.Print LastRow
'    Debug.Print 12
End Function

Function csv�ǂݍ���(Customer_name As Variant) As Variant
  Dim file As String, max_n As Long
  Dim buf As String, tmp As Variant, ary() As Variant
  Dim i As Long, n As Long, val As Long
  max_n = 0
  
  '����
  'file = "C:\test.csv" '�t�@�C���w��
  file = csv�t�@�C�����T��(Customer_name)
  
'  max_n = CreateObject("Scripting.FileSystemObject").OpenTextFile(file, 8).Line '�t�@�C���̍s���擾
  
  Open file For Input As #1 'CSV�t�@�C�����J��
        Do Until EOF(1)
            Line Input #1, buf
            max_n = max_n + 1
        Loop
  Close #1 'CSV�t�@�C�������
  
  ReDim ary(max_n - 1, 2) As Variant '�擾�����s����2�����z��̍Ē�`
    
  Open file For Input As #1 'CSV�t�@�C�����J��
      Do Until EOF(1) '�ŏI�s�܂Ń��[�v
      Line Input #1, buf '�ǂݍ��񂾃f�[�^��1�s���݂Ă���
      tmp = Split(buf, ",") '�J���}�ŕ���
      For i = 0 To UBound(tmp) '���ڐ��Ԃ񃋁[�v
        ary(n, i) = tmp(i) '�����������e��z��̍��ڂ֓����i0��ID, 1������, 2���l�j
      Next i
      n = n + 1 '�z��̎��̍s��
    Loop
  Close #1 'CSV�t�@�C�������
  
    csv�ǂݍ��� = ary
'
'  For i = 1 To UBound(ary)
'    Debug.Print ary(i, 0)
'  Next
End Function

Function csv�t�@�C�����T��(Customer_name As Variant) As Variant

    ship_date = Sheets("�󒍓���").Range("A1")  '�o�ד�
    
    '�p�X�̌���
    sFileFullPath = ThisWorkbook.Path
    For i = Len(sFileFullPath) To 0 Step -1
        If InStr(i, sFileFullPath, "\") > 0 Then
            '���݂̃t�H���_�����擾
            sFolderName = Mid(sFileFullPath, InStr(i, sFileFullPath, "\") + 1)
            '1��̊K�w�̃t�H���_�̂܂ł̃t���p�X���擾
            sParentFolderPath = Mid(sFileFullPath, 1, InStr(1, sFileFullPath, sFolderName) - 2)
            Exit For
        End If
    Next
    csvFilePath = sParentFolderPath & "\�s�b�L���O�\\csv\" & Customer_name & "\" & Year(ship_date) & "�N\" & Right("0" & Month(ship_date), 2) & "��"
    
    '�f�B���N�g�����݃`�F�b�N
    If Dir(csvFilePath, vbDirectory) = "" Then
        MsgBox "csv�f�B���N�g�������݂��܂���B" & vbCrLf & csvFilePath & vbCrLf & vbCrLf & "�y�o�ד��z" & ship_date & vbCrLf & "�y�o�א�z" & Customer_name & vbCrLf & "�@�s�b�L���O�\�̈�����s���Ă��Ȃ��\��������܂��B" & vbCrLf & "�A���Y�\A1�Z���̏o�ד����ԈႦ�Ă���\��������܂��B"
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
        MsgBox "csv�t�@�C������ł��B" & vbCrLf & csvFilePath & vbCrLf & vbCrLf & "�y�o�ד��z" & ship_date & vbCrLf & "�y�o�א�z" & Customer_name & vbCrLf & "�@�s�b�L���O�\�̈�����s���Ă��Ȃ��\��������܂��B" & vbCrLf & "�A���Y�\A1�Z���̏o�ד����ԈႦ�Ă���\��������܂��B"
        End
    End If
    
    ReDim filename(s - 1, 4) '�y0�z�t�@�C����.csv�y1�z�쐬���y2�z�N(�t�@�C����)�y3�z��0����(�t�@�C����)�y4�z��0����(�t�@�C����)
    
    With CreateObject("Scripting.FileSystemObject")
        For Each f In .GetFolder(csvFilePath).Files
            filename(cnt, 0) = f.Name
            filename(cnt, 1) = f.DateLastModified   '�쐬����f.DateCreated �X�V����f.DateLastModified  �A�N�Z�X����f.DateLastAccessed
            filename(cnt, 2) = Mid(filename(cnt, 0), 5, 4)  '�N �t�@�C����
            filename(cnt, 3) = Mid(filename(cnt, 0), 10, 2) '�� �t�@�C����
            filename(cnt, 4) = Mid(filename(cnt, 0), 13, 2) '�� �t�@�C����
            cnt = cnt + 1
        Next f
    End With
    
    Dim Max As Integer
    Max = 9999 '�����l��ݒ� ���L��for����9999�̂܂܂Ȃ�Acsv�f�[�^�ɊY���̏o�ד����Ȃ������Ƃ������ƂɂȂ�B
    For i = 0 To UBound(filename)
        If Str(Year(ship_date)) & Str(Right("0" & Month(ship_date), 2)) & Str(Right("0" & Day(ship_date), 2)) = Str(filename(i, 2)) & Str(filename(i, 3)) & Str(filename(i, 4)) Then
            If Max = 9999 Then
                Max = i
            End If
            If filename(i, 1) > filename(Max, 1) Then
                Max = i
            End If
        End If
    Next i
    
    'csv���݃`�F�b�N
    If Max = 9999 Then
        MsgBox "�Y���̏o�ד���csv�t�@�C��������܂���B" & vbCrLf & csvFilePath & vbCrLf & vbCrLf & "�y�o�ד��z" & ship_date & vbCrLf & "�y�o�א�z" & Customer_name & vbCrLf & "�@�s�b�L���O�\�̈�����s���Ă��Ȃ��\��������܂��B" & vbCrLf & "�A���Y�\A1�Z���̏o�ד����ԈႦ�Ă���\��������܂��B"
        End
    End If
    
    csv�t�@�C�����T�� = csvFilePath & "\" & filename(Max, 0)
    
End Function



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''�t�@�C���o�b�N�A�b�v''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Function �f�[�^�o�b�N�A�b�v(Customer_name As Variant)
    file_pass_to = �f�B���N�g���쐬�R�s�[��()
    
    '�R�s�[��̃p�X
    file_pass_to = file_pass_to & "\" & Customer_name & "_�s�b�L���O�\.xlsm"
    
    '�R�s�[���̃p�X
    file_pass_from = "\\Afnewt320-kyoyu\�Г����L\AFSKS\�s�b�L���O�\\" & Customer_name & "_�s�b�L���O�\.xlsm"
    
    '�t�@�C���̃R�s�[
    FileCopy file_pass_from, file_pass_to
    
End Function


Function �f�B���N�g���쐬�R�s�[��() As Variant
    Dim root As String
    Dim yyyy As String
    Dim mm As String
    Dim dd As String
    'A1�o�ד�
    Dim ship_date As Date
    
    ship_date = Workbooks("���Y�\.xlsm").Sheets("�󒍓���").Range("A1")

    ' root = ActiveWorkbook.Path & "\csv"
    root = "\\Afnewt320-kyoyu\�Г����L\AFSKS\�f�[�^�ۊ�"
    yyyy = Format(Year(ship_date), "0000�N")
    mm = Format(Month(ship_date), "00��")
    dd = Format(Day(ship_date), "00��")

    'F2 �o�א�
    Dim Customer_name As String
    Customer_name = Range("F2")
    
    Dim rtn As Long
    rtn = �f�B���N�g���쐬2(root, yyyy, mm, mm & dd)
    file_pass = root & "\" & yyyy & "\" & mm & "\" & mm & dd
    
    �f�B���N�g���쐬�R�s�[�� = file_pass
'    Select Case rtn
'        Case 0
'            MsgBox "�t�H���_���쐬���܂����B"
'        Case 1
'            MsgBox "�t�H���_�͑��݂��܂��B"
'        Case Else
'            MsgBox "�t�H���_�̍쐬�Ɏ��s���܂����B"
'    End Select
End Function

Function �f�B���N�g���쐬2(ParamArray arg()) As Long
    On Error GoTo ErrExit
    If Dir(Join(arg, "\"), vbDirectory) <> "" Then
        CreateDirectory = 1
        Exit Function
    End If
  
    Dim ary As Variant
    Dim i As Long
    For i = LBound(arg) To UBound(arg)
        ary = arg
        ReDim Preserve ary(i)
        If Dir(Join(ary, "\"), vbDirectory) = "" Then
            MkDir Join(ary, "\")
        End If
    Next
  
    CreateDirectory = 0
    Exit Function
  
ErrExit:
    CreateDirectory = 9
End Function

