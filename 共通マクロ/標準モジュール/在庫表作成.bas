Attribute VB_Name = "�݌ɕ\�쐬"
Sub �����݌ɕ\�쐬()
    Call �����G�N�Z���쐬
    Call �����u��.�u��main
    Call �t�@�C���ۑ�(dateList, fileList)
    Call �ی�.�ی샍�b�N�S�V�[�g�z��
    Call ��ʂƃA���[�g�\��
    Call �I��
End Sub

Sub �����G�N�Z���쐬()
'�Z�b�e�B���O �v�v�v�v�v
    Set dateList = �݌ɕ\�쐬_�u��.���t�N���X()
    Set fileList = �݌ɕ\�쐬_�u��.�t�@�C���N���X()
    Call �쐬�m�F
    Call �O����(dateList, fileList)
    Call �t�@�C�����݊m�F(dateList, fileList)
    Call �t�@�C���쐬(dateList, fileList)
    Call ���v���z�V�[�gB2�����֕ύX(dateList, fileList)
'    Call �݌ɕ\�쐬_�u��.main
'    Call �t�@�C���ۑ�(dateList, fileList)
'    Call �ی�.�ی샍�b�N�S�V�[�g�z��
'    Call ��ʂƃA���[�g�\��
'    Call �I��
End Sub



Sub �쐬�m�F()
    boool = MsgBox("�����݌ɕ\�̍쐬�@���s���܂����A��낵���ł����H", vbYesNo + vbQuestion)
    If boool = vbYes Then
        Exit Sub
    Else
        End
    End If
End Sub

Sub �O����(dateList As Variant, fileList As Variant)
    Call ��ʂƃA���[�g��\��
    Call �f�[�^�`�F�b�N(dateList, fileList)
End Sub

Sub ��ʂƃA���[�g��\��()
    Application.DisplayAlerts = False
    Application.ScreenUpdating = False
End Sub

Sub ��ʂƃA���[�g�\��()
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
End Sub

Sub �f�[�^�`�F�b�N(dateList As Variant, fileList As Variant)
   Debug.Print "�udateList.date_now�v " & dateList.date_now
   Debug.Print "�udateList.date_next�v " & dateList.date_next
   Debug.Print "�udateList.date_last�v " & dateList.date_last
   Debug.Print "�udateList.H_AL_nowMonth�v " & dateList.H_AL_nowMonth
   Debug.Print "�udateList.H_AL_nextMonth�v " & dateList.H_AL_nextMonth
   Debug.Print "�udateList.G_now_date�v " & dateList.G_now_date
   Debug.Print "�udateList.G_last_date�v " & dateList.G_last_date
   Debug.Print "�udateList.G_next_date�v " & dateList.G_next_date
   Debug.Print "�udateList.kowake_now_date�v " & dateList.kowake_now_date
   Debug.Print "�udateList.kowake_last_date�v " & dateList.kowake_last_date
   Debug.Print "�udateList.kowake_next_date�v " & dateList.kowake_next_date

   Debug.Print "�ufileList.this_filename�v " & fileList.this_filename
   Debug.Print "�ufileList.mypath�v " & fileList.mypath
   Debug.Print "�ufileList.mybook�v " & fileList.mybook
   Debug.Print "�ufileList.mybook_month�v " & fileList.mybook_month
   Debug.Print "�ufileList.mmfn�v " & fileList.mmfn
   Debug.Print "�ufileList.fn�v " & fileList.fn
End Sub

Sub ���v���z�V�[�gB2�����֕ύX(dateList As Variant, fileList As Variant)
    ActiveWorkbook.Sheets("���v���z").Activate
    Range("B2").Formula = dateList.date_next '���v���z�V�[�gB2�����֕ύX
End Sub


Sub �t�@�C�����݊m�F(dateList As Variant, fileList As Variant)
   If Dir(fileList.next_mybook) = "" Then
        'MsgBox "�����݌ɕ\���쐬���܂�"
    Else
        MsgBox "�������̍݌ɕ\��" & vbNewLine & _
               "���ɑ��݂��Ă��܂��B" & vbNewLine & _
               "  " & vbNewLine & _
               "�����𒆎~���܂��B"
        End
    End If
End Sub

Sub �t�@�C���쐬(dateList As Variant, fileList As Variant)

    Dim mybk As Workbook
    Set mybk = ThisWorkbook
    mybk.SaveAs (fileList.next_mybook)

    MsgBox fileList.next_mybook & "�쐬���܂���"
End Sub

Sub �t�@�C���ۑ�(dateList As Variant, fileList As Variant)
        ActiveWorkbook.Save
End Sub


Sub �I��()
    MsgBox "�I�����܂����B"
End Sub
