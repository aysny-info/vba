Function csv�ǂݍ���() As Variant
  Dim file As String, max_n As Long
  Dim buf As String, tmp As Variant, ary() As Variant
  Dim i As Long, n As Long, val As Long
  max_n = 0

  '����
  'file = "C:\test.csv" '�t�@�C���w��
  file = "\\192.168.100.105\�Vrev_files\order_YOK.csv"

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