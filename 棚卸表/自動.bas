Attribute VB_Name = "����"
Sub ����()
        ThisWorkbook.Activate
        Call �ی�.�S�ی����
        Worksheets("�ݒ�").Range("X1").Value = Now
End Sub

Sub �΂���������()
    On Error GoTo Error1 '�ꉞ�G���[�΍�
    Dim res As Integer
    Dim CopyBook As String
    Dim hiduke As String
    hiduke = Format(Month(Now), "00") & Format(Day(Now), "00") & "_" & Format(Hour(Now), "00") & Format(Minute(Now), "00") & Format(Second(Now), "00")
        
    'Backup��̕ۑ��恕�u�b�N��
    CopyBook = ThisWorkbook.Path & "\�o�b�N�A�b�v\" & "BackupFile" & hiduke & ".xlsm"
    ActiveWorkbook.SaveCopyAs CopyBook  'Backup�ۑ�����R�[�h

    Exit Sub
Error1:         '�G���[�����������ꍇ�͂����֔��
    MsgBox "�G���[�ԍ�:" & Err.Number & vbLf & _
    "�G���[���e�F" & Err.Description & vbLf
    Exit Sub
End Sub

Sub ���t�̓���()
    Worksheets("�ݒ�").Range("A1").Value = Format(Year(Now), "00")
    Worksheets("�ݒ�").Range("C1").Value = Format(Month(Now), "00")
    Worksheets("�ݒ�").Range("E1").Value = Format(Day(Now), "00")
End Sub

Sub �g�p��()
    Dim use_file_name(2) As String
    use_file_name(1) = "�g�p��.txt"
    use_file_name(2) = "�󂢂Ă܂�.txt"
    
    Dim FSO As Object
    Set FSO = CreateObject("Scripting.FileSystemObject")
    On Error GoTo myError
    FSO.GetFile(ThisWorkbook.Path & "\" & use_file_name(2)).Name = use_file_name(1)
    Exit Sub
    Set FSO = Nothing
    
myError:
    MsgBox "�u�g�p���v�u �󂢂Ă܂��v ���@�\���Ă��Ȃ��\������(�������ėǂ�)"
End Sub

Sub �󂢂Ă܂�()
    Dim use_file_name(2) As String
    use_file_name(1) = "�g�p��.txt"
    use_file_name(2) = "�󂢂Ă܂�.txt"
    
    Dim FSO As Object
    Set FSO = CreateObject("Scripting.FileSystemObject")
    On Error GoTo myError
    FSO.GetFile(ThisWorkbook.Path & "\" & use_file_name(1)).Name = use_file_name(2)
    Exit Sub
    Set FSO = Nothing
    
myError:
    MsgBox "�u�g�p���v�u �󂢂Ă܂��v ���@�\���Ă��Ȃ��\������(�������ėǂ�)"
End Sub

