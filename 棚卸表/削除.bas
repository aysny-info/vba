Attribute VB_Name = "�폜"

''''''''''''''''''''''''''''''''               �폜�֘A         ''''''''''''''''''''''''''''''''''''''
Sub �폜_�d�|��()
     Worksheets("�d�|").Activate
     Worksheets("�d�|").Range(Cells(11, 1), Cells(9010, 34)).Clear
End Sub

Sub �폜_����()
     Worksheets("����").Activate
     Worksheets("����").Range(Cells(11, 1), Cells(9010, 34)).Clear
End Sub

Sub �폜_�݌ɐ�()
     Worksheets("�݌ɐ�").Activate
     Worksheets("�݌ɐ�").Range(Cells(11, 1), Cells(9010, 47)).Clear
End Sub

Sub �폜_CN����()
     Worksheets("CN����").Activate
     Worksheets("CN����").Range(Cells(11, 15), Cells(9010, 47)).Clear
     Worksheets("CN����").Range(Cells(11, 2), Cells(9010, 2)).Clear
                   '����CD1001-9999
              Dim i As Long, B As Variant
              ReDim B(9010, 0)
              For i = 0 To 8998
                B(i, 0) = i + 1001
              Next i
              Range("M11:M9999") = B
End Sub

Sub �폜_IY����()
     Worksheets("IY����").Activate
     Worksheets("IY����").Range(Cells(11, 3), Cells(9010, 34)).Clear
              '����CD1001-9999
              Dim i As Long, B As Variant
              ReDim B(9010, 0)
              For i = 0 To 8998
                B(i, 0) = i + 1001
              Next i
              Range("A11:A9999") = B
End Sub

Sub �폜_ALL()
    Call �폜_�d�|��
    Call �폜_����
    Call �폜_�݌ɐ�
    Call �폜_CN����
    Call �폜_IY����
End Sub
