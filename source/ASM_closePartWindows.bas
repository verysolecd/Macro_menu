Attribute VB_Name = "ASM_closePartWindows"
'Attribute VB_Name = "m10_closePartWindows"
' Part窗口一次性全关闭
'{GP:3}
'{EP:CLSpart}
'{Caption:关闭零件窗}
'{ControlTipText: 点击后一次性全关闭所有零件窗口}
'{背景颜色: 12648447}

Sub CLSpart()
Dim wds, WD
 On Error Resume Next
   wds = CATIA.Windows
    If wds.count <= 1 Then
           MsgBox "没有打开的零件窗口"
           Exit Sub
    End If
    For i = 1 To wds.count
        Set WD = wds.item(i)
        If KCL.isobjtype(WD.Parent, "PartDocument") Then
            WD.Close
        End If
    Next
    On Error GoTo 0
End Sub

'
