Attribute VB_Name = "OTH_Color"
'Attribute VB_Name = "OTH_Color"
'{GP:6}
'{Ep:CATmain}
'{Caption:背景颜色}
'{ControlTipText:白黑色背景切换}
'{BackColor: }
' 更新按钮文字的公共函数


Sub CATMain()
    On Error GoTo errorhandler
    If CATIA.Windows.count < 1 Then
        MsgBox "没有打开的窗口"
        Exit Sub
    End If
        
    Dim oWindow, oViewer
    Set oWindow = CATIA.ActiveWindow
    Set oViewer = oWindow.ActiveViewer
    
    oWindow.Layout = catWindowGeomOnly
    oViewer.Reframe
    Dim MyViewer: Set MyViewer = CATIA.ActiveWindow.ActiveViewer
    Dim currentColor(2)
    MyViewer.GetBackgroundColor currentColor
    ' 根据当前背景色直接切换
    If currentColor(0) = 1 And currentColor(1) = 1 And currentColor(2) = 1 Then
        ' 当前是白色背景，切换到默认背景
        MyViewer.PutBackgroundColor Array(0.2, 0.2, 0.4)
        oWindow.Layout = catWindowSpecsAndGeom
    Else
        ' 当前是默认背景，切换到白色背景
        MyViewer.PutBackgroundColor Array(1, 1, 1)
        oWindow.Layout = catWindowGeomOnly
    End If
    
    On Error GoTo 0
    
errorhandler:
    If Err.Number <> 0 Then
        Select Case Err.Number
            Case 1001
                MsgBox "CATIA 错误：" & Err.Description, vbCritical
                Err.Clear
                Exit Sub
            Case 1002
        End Select
    End If
End Sub
