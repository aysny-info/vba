Attribute VB_Name = "�݌ɕ\�쐬_�G�N�Z��"
Sub �쐬main()
'�Z�b�e�B���O �v�v�v�v�v
    Set dateList = �݌ɕ\�쐬_�u��.���t�N���X()
    Set fileList = �݌ɕ\�쐬_�u��.�t�@�C���N���X()

    Call ��ʂƃA���[�g��\��
    Call �݌ɕ\�쐬_�u��.�f�[�^�`�F�b�N(dateList, fileList)
    Call �t�@�C�����݊m�F(dateList, fileList)
    Call �t�@�C���쐬(dateList, fileList)
    Call �݌ɕ\�쐬_�u��.main
    Call �t�@�C���ۑ�(dateList, fileList)
    Call ��ʂƃA���[�g�\��
    Call �I��
End Sub

Sub �t�@�C�����݊m�F(dateList As Variant, fileList As Variant)
   If Dir(fileList.next_mybook) = "" Then
        MsgBox "�����݌ɕ\���쐬���܂�"
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
'    boool = MsgBox("�t�@�C����ۑ����܂����H", vbYesNo + vbQuestion)
'    If boool = vbYes Then
'        ActiveWorkbook.Save
'    Else
'        End
'    End If
        ActiveWorkbook.Save
End Sub

Sub ��ʂƃA���[�g��\��()
    Application.DisplayAlerts = False
    Application.ScreenUpdating = False
End Sub

Sub ��ʂƃA���[�g�\��()
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
End Sub

Sub �I��()
    MsgBox "�I�����܂����B"
End Sub
