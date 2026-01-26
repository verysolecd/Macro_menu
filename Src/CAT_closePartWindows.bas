Attribute VB_Name = "CAT_closePartWindows"
' Part窗口一次性全关闭
'{GP:7}
'{EP:CLSpart}
'{Caption:关闭零件窗}
'{ControlTipText: 点击后一次性全关闭所有零件窗口}
'{背景颜色: 12648447}

Private Const mdlname As String = "CAT_closePartWindows"
Sub CLSpart()
Dim wds, WD
 On Error Resume Next
   Set wds = CATIA.Windows
    If wds.count <= 1 Then
           MsgBox "没有打开的零件窗口"
           Exit Sub
    End If
    For i = 1 To wds.count
        Set WD = wds.item(i)
        If KCL.IsObj_T(WD.Parent, "PartDocument") Then
            WD.Close
        End If
    Next
    On Error GoTo 0
End Sub

'
