Sub ReplaceHyperlinks()
'Updateby Extendoffice
Dim Ws As Worksheet
Dim xHyperlink As Hyperlink
Dim xOld As String, xNew As String
xTitleId = "KutoolsforExcel"
Set Ws = Application.ActiveSheet
'xOld = Application.InputBox("..\..\", xTitleId, "", Type:=2)
'xNew = Application.InputBox("\\Afnewt320-kyoyu\", xTitleId, "", Type:=2)

xOld = "\\Afnewt320-kyoyu\�Г����L\�y�c�ƕ��z\�y�V�E�H��A���z\"
xNew = "\\\\Afnewt320-kyoyu\\�Г����L\\"


Application.ScreenUpdating = False
For Each xHyperlink In Ws.Hyperlinks
    xHyperlink.Address = Replace(xHyperlink.Address, xOld, xNew)
    Debug.Print xHyperlink.Address
Next
Application.ScreenUpdating = True



End Sub
