Attribute VB_Name = "ASM_closePartWindows"
'Attribute VB_Name = "m10_closePartWindows"
' Part窗口一次性全关闭
'{GP:3}
'{EP:CLSpart}
'{Caption:关闭零件窗}
'{ControlTipText: 点击后一次性全关闭所有零件窗口}
'{背景颜色: 12648447}

Sub CLSpart()
Dim CATIA, wds, wd
 On Error Resume Next
    Set CATIA = GetObject(, "CATIA.Application")
    Set wds = CATIA.Windows
    If wds.Count < 1 Then
           MsgBox "没有打开的窗口"
           Exit Sub
    End If
    For i = 1 To wds.Count
        Set wd = wds.item(i)
        If KCL.IsType_Of_T(wd.Parent, "PartDocument") Then
            wd.Close
        End If
    Next
    On Error GoTo 0
End Sub

'
