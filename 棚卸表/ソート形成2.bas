Attribute VB_Name = "�\�[�g�`��2"
Sub �\�[�g�`��2_���V�s()
    With ActiveSheet
        Range("���V�s[[#Headers],[�d���於]]").Select
        If .FilterMode Then .ShowAllData
    End With

    ActiveWorkbook.Worksheets("�`��2").ListObjects("���V�s").Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("�`��2").ListObjects("���V�s").Sort.SortFields.Add Key _
        :=Range("���V�s[[#All],[�d���於]]"), SortOn:=xlSortOnValues, Order:= _
        xlDescending, DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets("�`��2").ListObjects("���V�s").Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
End Sub

Sub �\�[�g�`��2�N���A()
    ActiveWorkbook.Worksheets("�`��2").ListObjects("���V�s").Sort.SortFields.Clear
    ActiveSheet.Range("B2").Select
End Sub


