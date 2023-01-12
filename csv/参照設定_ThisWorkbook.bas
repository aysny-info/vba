Private Sub Workbook_Open()
    Application.Calculation = xlCalculationManual   '手動計算
    Application.ScreenUpdating = False

    '参照設定の追加
    Call Microsoft_Scripting_Runtime
    Call microsoft_visual_basic_for_applications_extensibility_5_3
    Call microsoft_activex_data_objects_2_8_library
    
    Application.Calculation = xlAutomatic   '自動計算
    Application.Calculate                   '再計算
    Application.ScreenUpdating = True
End Sub

