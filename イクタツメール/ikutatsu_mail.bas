Sub �C�N�^�c����()
        ActiveWorkbook.Sheets("�ꗗ�A").Activate            '�A�N�e�B�u
        Application.Calculation = xlCalculationAutomatic  '�����v�Z
        Application.Calculate                             '�Čv�Z
        '�ŏI�s�擾
        Dim LastRow As Long
        LastRow = Cells(Rows.Count, 11).End(xlUp).Row
        '������z��֊i�[     F�񂩂�AJ��
        tikan_cell = Range(Cells(1, 11), Cells(LastRow, 27))

        Dim judgeIkutatsu, sijisuu As Integer
        judgeIkutatsu = 0
        sijisuu = 0
        '�u��
        For s = 1 To (LastRow - 1)
                    If tikan_cell(s, 1) Like "*�C�N�^�c*" Then
                        judgeIkutatsu = 1
                        sijisuu = sijisuu + tikan_cell(s, 16)
                    End If
        Next s

        Debug.Print judgeIkutatsu
        Debug.Print sijisuu

        '���[��
        Call ���[�����M(sijisuu)
End Sub

Public Function ���[�����M(sijisuu As Integer)
    Set oApp = CreateObject("Outlook.Application")
    Set myNameSpace = oApp.GetNamespace("MAPI")
    Set myFolder = myNameSpace.GetDefaultFolder(6)
    'myFolder.Display 'OUTLOOK�N��
    Set objMail = oApp.CreateItem(olMailItem)

    '���[���N���X
    Dim mailList  As mailClass
    Set mailList = New mailClass

    '���[��
    mailList.meado = ���[�����M���X�g()
    mailList.kenmei = "����߂�"
    mailList.naiyou = Worksheets("���[����").Range("B2").Value + vbCrLf + vbCrLf + Worksheets("���[����").Range("B3").Value + vbCrLf + vbCrLf + "���s���� : " + Str(sijisuu / 10) + "�� + �\����" + vbCrLf + " (�����p�b�N�� " + Str(sijisuu) + "pc)"

    '���[���֔��f
    With objMail
        .To = mailList.meado(0)
        .CC = mailList.meado(1)
        .Subject = mailList.kenmei
        .Body = mailList.naiyou
        .BodyFormat = 2
        .Display            'OUTLOOK���M��ʂ̋N��
    End With

End Function



Public Function ���[�����M���X�g() As Variant
    Worksheets("���X�g").Activate
    Dim meadoSeizoTemp As Range, meadoSystemTemp As Range, meadoEigyoTemp As Range, meadoKanriTemp As Range

    '�u���X�g�v�V�[�g�̈�ԉ�
    Dim Last_Row_Seizo As Long, Last_Row_System As Long, Last_Row_kanri As Long
    Last_Row_Seizo = Worksheets("���X�g").Cells(Rows.Count, 5).End(xlUp).Row
    Last_Row_System = Worksheets("���X�g").Cells(Rows.Count, 7).End(xlUp).Row
    Last_Row_kanri = Worksheets("���X�g").Cells(Rows.Count, 14).End(xlUp).Row

    '���A�h���ꂼ��擾
    Set meadoSeizoTemp = Worksheets("���X�g").Range(Cells(4, 5), Cells(Last_Row_Seizo, 5))
    Set meadoSystemTemp = Worksheets("���X�g").Range(Cells(4, 7), Cells(Last_Row_System, 7))
    Set meadoKanriTemp = Worksheets("���X�g").Range(Cells(4, 14), Cells(Last_Row_kanri, 14))

    '���A�h���ꂼ��P���ɐ��`
    Dim meadoSeizo As String, meadoSystem As String, meadoEigyo As String, meadoKanri As String
    Dim i As Long
            '�������A�h
            For i = 0 To UBound(meadoSeizoTemp(), 1)
                If meadoSeizoTemp(i) = "" Then

                Else
                    meadoSeizo = meadoSeizo + meadoSeizoTemp(i)
                    meadoSeizo = meadoSeizo + ";"
                End If
            Next i

            '�V�X�e�����A�h
            For i = 0 To UBound(meadoSystemTemp(), 1)
                If meadoSystemTemp(i) = "" Then

                Else
                    meadoSystem = meadoSystem + meadoSystemTemp(i)
                    meadoSystem = meadoSystem + ";"
                End If
            Next i

            '�Ǘ����A�h
            For i = 0 To UBound(meadoKanriTemp(), 1)
                If meadoKanriTemp(i) = "" Then

                Else
                    meadoKanri = meadoKanri + meadoKanriTemp(i)
                    meadoKanri = meadoKanri + ";"
                End If
            Next i

    '���M�惁�A�h�܂Ƃ�
    Dim meadoMatome(1) As String
    '����
        meadoMatome(0) = meadoMatome(0) + meadoKanri
    'CC
        meadoMatome(1) = meadoMatome(1) + meadoSeizo
        meadoMatome(1) = meadoMatome(1) + meadoSystem

    ���[�����M���X�g = meadoMatome

End Function
