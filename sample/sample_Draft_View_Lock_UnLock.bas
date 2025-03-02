'Attribute VB_Name = "sample_Draft_View_Lock_UnLock"
' 图纸视图的锁定与解锁
'{GP:21}
'{EP:CATMain}
'{Caption:锁定_解锁}
'{ControlTipText: 可以进行图纸视图的锁定与解锁}
'{BackColor:12648447}

Option Explicit
Sub CATMain()
    ' 检查是否可以执行
    If Not CanExecute("DrawingDocument") Then
     Exit Sub
    Dim Views As DrawingViews
    Set Views = CATIA.ActiveDocument.Sheets.ActiveSheet.Views
    If Views.Count < 3 Then
     Exit Sub
    Dim View As DrawingView
    Set View = Views.Item(3)
    Dim LockState As Boolean
    LockState = View.LockStatus
    Dim Msg As String
    If LockState Then
        Msg = "解锁"
        LockState = False
    Else
        Msg = "锁定"
        LockState = True
    End If
    Dim i As Long
    For i = 3 To Views.Count
        Set View = Views.Item(i)
        View.LockStatus = LockState
    Next
    MsgBox "视图已成功" & Msg & "。"
End Sub
