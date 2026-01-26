Attribute VB_Name = "CAT_SWScr"
'{GP:7}
'{Ep:switchRefresh}
'{Caption: 屏幕更新}
'{ControlTipText:禁止屏幕更新以防止卡顿}
'{BackColor: }
Private Const mdlname As String = "CAT_SWScr"
Sub switchRefresh()
On Error Resume Next
    CATIA.ActiveWindow.ActiveViewer.Update
    On Error GoTo 0
End Sub
Sub isRefresh()
 istrue = CATIA.RefreshDisplay
 MsgBox istrue
End Sub
