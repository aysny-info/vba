Public Function ���[�����M�蓮(sijisuu As Integer, seizoubi As Date, shikabi As Date)
    Set oApp = CreateObject("Outlook.Application")
    Set myNameSpace = oApp.GetNamespace("MAPI")
    Set myFolder = myNameSpace.GetDefaultFolder(6)
    'myFolder.Display 'OUTLOOK�N��
    Set objMail = oApp.CreateItem(olMailItem)

    '���[��
    meado = "hashimoto@aysny.co.jp; oba@aysny.co.jp;"
    kenmei = "�y�X�V�ʒm�z�R�[�v�f���t���[�Y��"
    naiyou = "�����l�ł��B" + vbCrLf + "�X�V���������m�点�v���܂��B"

    '���[���֔��f
    With objMail
        .To = meado
        '.CC = mailList.meado(1)
        .Subject = kenmei
        .Body = naiyou
        .BodyFormat = 2
        .Display            'OUTLOOK���M��ʂ̋N��
    End With

End Function