    Dim rc As VbMsgBoxResult
    rc = MsgBox("���Z�b�g���܂����H" & vbCrLf & vbCrLf & "�@�o�ד��𗂓���" & vbCrLf & "�Acsv�V�[�g�̃��Z�b�g", vbYesNo + vbQuestion)
    If rc = vbNo Then
        MsgBox "�����𒆎~���܂�", vbCritical
        Exit Sub
    End If
