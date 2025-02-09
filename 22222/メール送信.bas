Sub Mailbox_Open()
    Application.Calculation = xlCalculationManual
    Application.ScreenUpdating = False
    'アクティブ
    Workbooks(ThisWorkbook.Name).Worksheets("更新日").Activate
    Workbooks(ThisWorkbook.Name).Worksheets("更新日").Range("A1").Select
       
    Customer_name = Mid(ThisWorkbook.Name, InStrRev(ThisWorkbook.Name, "\") + 1, InStrRev(ThisWorkbook.Name, ".") - InStrRev(ThisWorkbook.Name, "\") - 1)
    Cells(3, 3) = Now   'C3 集計表用に更新日を入力
    'Cells(5 + Month(Now), 4) = Now  '各月の更新日を入力
    ThisWorkbook.Save
    Call バックアップ_main
    Call 集計表の更新
'    Call シート抽出
    Call メール作成(Customer_name)
'    Call testXXX
    ThisWorkbook.Activate
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    Application.Calculate
    
    ThisWorkbook.Save
End Sub

'Sub シート抽出()
'    Dim mm As Integer
'    Dim ws As Worksheet
'
'    '今月
'    mm1 = Trim(Month(Date))
'    '翌月
'    mm2 = Trim(Month(DateAdd("m", 1, Date)))
'    '翌々月
'    mm3 = Trim(Month(DateAdd("m", 2, Date)))
'    Debug.Print mm1
'    Debug.Print mm2
'    Debug.Print mm3
'
'    Sheets(Array(mm1 & "月", mm2 & "月", mm3 & "月")).Select
'    Sheets(Array(mm1 & "月", mm2 & "月", mm3 & "月")).Copy
'
'
'    Customer_name = Mid(ThisWorkbook.Name, InStrRev(ThisWorkbook.Name, "\") + 1, InStrRev(ThisWorkbook.Name, ".") - InStrRev(ThisWorkbook.Name, "\") - 1)
'    ActiveWorkbook.SaveAs "\\Afnewt320-kyoyu\社内共有\【生産管理】\販売計画\mail\" & Customer_name & ".xlsx"
'
'    Sheets(mm1 & "月").Select
'
'    Call 保護.保護_全解除
'    Sheets(Array(mm1 & "月", mm2 & "月", mm3 & "月")).Select
'    Cells.Select
'    Selection.Copy
'    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
'    :=False, Transpose:=False
'
'
''    複数シート選択解除
'    ActiveWindow.SelectedSheets(1).Select
'    ActiveWorkbook.Save
'    ActiveWorkbook.Close
'    Sheets("更新日").Select
'
'End Sub

Function メール作成(ship_name As Variant)
    '--- Outlook操作のオブジェクト ---'
    Dim objOutlook As Object
    Set objOutlook = CreateObject("Outlook.Application")
    
    '--- メールオブジェクト ---'
    Dim objMail As Object
    Set objMail = objOutlook.CreateItem(0)
        
    '--- メールの内容を格納する変数 ---'
    Dim toStr As String
    Dim ccStr As String
    Dim bccStr As String
    Dim subjectStr As String
    Dim bodyStr As String
    
    '--- 宛先の内容 ---'
    toStr = Worksheets("更新日").Range("G3").Value '"[宛先のメールアドレス]"
    ccStr = "" '"[CCのメールアドレス]"
    bccStr = "" ' "[BCCのメールアドレス]"
    
    '--- 件名の内容 ---'
    subjectStr = "販売計画更新通知_" + ship_name
    
    '--- 本文の内容 ---'
    Dim ship_date As Date
    
    ship_date = Now
    
    yyyy = Format(Year(ship_date), "0000年")
    mm = Format(Month(ship_date), "00月")
    dd = Format(Day(ship_date), "00日")
    hh = Format(Hour(ship_date), "00時")
    
    ship_filename = "【" & yyyy & mm & dd & hh & "】"
    
    linkA = "\\Afnewt320-kyoyu\社内共有\【生産管理】\販売計画\データ保管" & "\" & ship_name & "\" & yyyy & "\" & mm & "\" & ship_filename & ship_name & ".xlsm"
    
    bodyStr = "<div>お疲れ様です。</div> " + "<br>" + "<div>販売計画表を更新しました。</div>" + "<br>" + "<div><a href=""" & linkA & """> " + linkA + "</a></div>"
        
    '--- 条件を設定 ---'
    objMail.To = toStr
    objMail.CC = ccStr
    objMail.BCC = bccStr
    objMail.Subject = subjectStr
    'objMail.BodyFormat = olFormatPlain
    objMail.HTMLBody = bodyStr
    
    '--- 添付ファイルのパス ---'
    Dim attachmentPath As String
'    attachmentPath = "\\Afnewt320-kyoyu\社内共有\【生産管理】\販売計画\" + ship_name + ".xlsm"
'    attachmentPath = "\\Afnewt320-kyoyu\社内共有\【生産管理】\販売計画\mail\" & ship_name & ".xlsx"
    
    '--- 添付ファイルを設定 ---'
'    Call objMail.Attachments.Add(attachmentPath)
    
    '--- メールを表示 ---'
    objMail.Display
    
    '--- メールを送付 ---'
    'objMail.Send

End Function



'Sub testXXX()
'On Error GoTo ErrorHandler
'    Dim filePath As String
'    filePath = "\\Afnewt320-kyoyu\社内共有\【生産管理】\販売計画\mail"
'
'    Kill filePath & "\" & "*.xlsx"
'
'ErrorHandler:
'    'エラー処理
'    If Err.Number <> 0 Then
'    End If
'End Sub


