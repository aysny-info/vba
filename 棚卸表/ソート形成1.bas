Attribute VB_Name = "�\�[�g�`��1"
Sub �\�[�g�`��1_�V��()
    With ActiveSheet
        Range("�V��[[#Headers],[�d���於]]").Select
        If .FilterMode Then .ShowAllData
    End With
    
    ActiveWorkbook.Worksheets("�`��1").ListObjects("�V��").Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("�`��1").ListObjects("�V��").Sort.SortFields.Add Key _
        :=Range("�V��[[#All],[�d���於]]"), SortOn:=xlSortOnValues, Order:= _
        xlDescending, DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets("�`��1").ListObjects("�V��").Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
End Sub

Sub �\�[�g�`��1_���i�Ǘ�()
    With ActiveSheet
        Range("���i�Ǘ�[[#Headers],[�d���於]]").Select
        If .FilterMode Then .ShowAllData
    End With
    
    ActiveWorkbook.Worksheets("�`��1").ListObjects("���i�Ǘ�").Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("�`��1").ListObjects("���i�Ǘ�").Sort.SortFields.Add Key _
        :=Range("���i�Ǘ�[[#All],[�d���於]]"), SortOn:=xlSortOnValues, Order:= _
        xlDescending, DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets("�`��1").ListObjects("���i�Ǘ�").Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
End Sub

Sub �\�[�g�`��1_�①��()
    With ActiveSheet
        Range("�①��[[#Headers],[�d���於]]").Select
        If .FilterMode Then .ShowAllData
    End With
    
    ActiveWorkbook.Worksheets("�`��1").ListObjects("�①��").Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("�`��1").ListObjects("�①��").Sort.SortFields.Add Key _
        :=Range("�①��[[#All],[�d���於]]"), SortOn:=xlSortOnValues, Order:= _
        xlDescending, DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets("�`��1").ListObjects("�①��").Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
End Sub

Sub �\�[�g�`��1_�Ⓚ��()
    With ActiveSheet
        Range("�Ⓚ��[[#Headers],[�d���於]]").Select
        If .FilterMode Then .ShowAllData
    End With
    
    ActiveWorkbook.Worksheets("�`��1").ListObjects("�Ⓚ��").Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("�`��1").ListObjects("�Ⓚ��").Sort.SortFields.Add Key _
        :=Range("�Ⓚ��[[#All],[�d���於]]"), SortOn:=xlSortOnValues, Order:= _
        xlDescending, DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets("�`��1").ListObjects("�Ⓚ��").Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
End Sub

Sub �\�[�g�`��1_���̑�()
    With ActiveSheet
        Range("���̑�[[#Headers],[�d���於]]").Select
        If .FilterMode Then .ShowAllData
    End With
    
    ActiveWorkbook.Worksheets("�`��1").ListObjects("���̑�").Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("�`��1").ListObjects("���̑�").Sort.SortFields.Add Key _
        :=Range("���̑�[[#All],[�d���於]]"), SortOn:=xlSortOnValues, Order:= _
        xlDescending, DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets("�`��1").ListObjects("���̑�").Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
End Sub

Sub �\�[�g�`��1_ALL()
    Call �\�[�g�`��1_�V��
    Call �\�[�g�`��1_���i�Ǘ�
    Call �\�[�g�`��1_�①��
    Call �\�[�g�`��1_�Ⓚ��
    Call �\�[�g�`��1_���̑�
    
    ActiveSheet.Range("M1").Select
End Sub

Sub �\�[�g�`��1�N���A_ALL()
    ActiveWorkbook.Worksheets("�`��1").ListObjects("�V��").Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("�`��1").ListObjects("���i�Ǘ�").Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("�`��1").ListObjects("�①��").Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("�`��1").ListObjects("�Ⓚ��").Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("�`��1").ListObjects("���̑�").Sort.SortFields.Clear
    
    ActiveSheet.Range("P1").Select
End Sub


