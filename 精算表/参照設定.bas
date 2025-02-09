Sub 使用者の参照設定のGUIDを調べる()
    'https://kouten0430.hatenablog.com/entry/2017/10/22/134852
    Dim myRef As Variant
    Dim i As Integer
        i = 1
    For Each myRef In ActiveWorkbook.VBProject.References
        i = i + 1
        Debug.Print (myRef.Name)
        Debug.Print (myRef.GUID)
        Debug.Print (myRef.Major)
        Debug.Print (myRef.Minor)
    Next
End Sub

Sub Microsoft_Scripting_Runtime()

    On Error GoTo Err
    
'    Temp = "{420B2830-E718-11CF-893D-00A0C9054228}"
'    Set ref = Application.VBE.ActiveVBProject.References.AddFromGuid(Temp, 1, 0)

    '********************** Microsoft Scripting RuntimeのGUID ********************
    Const MSR_GUID = "{420B2830-E718-11CF-893D-00A0C9054228}"
    Application.VBE.ActiveVBProject.References.AddFromGuid MSR_GUID, 1, 0   '参照設定を追加
    
    Debug.Print "参照設定を追加しました！" & MSR_GUID
    
    Exit Sub
    
Err:
    Debug.Print "エラーが発生しました！" & vbCrLf & Err.Description
 
End Sub

Sub microsoft_visual_basic_for_applications_extensibility_5_3()

    On Error GoTo Err
    
    '********************** microsoft visual basic for applications extensibility 5.3のGUID ********************
    Const MSR_GUID = "{0002E157-0000-0000-C000-000000000046}"
    Application.VBE.ActiveVBProject.References.AddFromGuid MSR_GUID, 1, 0   '参照設定を追加
    
    Debug.Print "参照設定を追加しました！" & MSR_GUID
    
    Exit Sub
    
Err:
    Debug.Print "エラーが発生しました！" & vbCrLf & Err.Description
 
End Sub

Sub microsoft_activex_data_objects_2_8_library()

    On Error GoTo Err
    
    '********************** microsoft activex data objects 2.8 libraryのGUID ********************
    Const MSR_GUID = "{2A75196C-D9EB-4129-B803-931327F72D5C}"
    Application.VBE.ActiveVBProject.References.AddFromGuid MSR_GUID, 1, 0   '参照設定を追加
    
    Debug.Print "参照設定を追加しました！" & MSR_GUID
    
    Exit Sub
    
Err:
    Debug.Print "エラーが発生しました！" & vbCrLf & Err.Description
 
End Sub

