Attribute VB_Name = "M60_Color"
'{GP:6}
'{Ep:CATmain}
'{Caption:白色背景}
'{ControlTipText:改白色背景好截图}
'{BackColor: }

Const CF_BITMAP = 2
Sub CATMain()
 On Error GoTo ErrorHandler
        If CATIA.Documents.Count = 0 Then
            Err.Raise 1001, , "未检测到打开的CATIA文档"
            Exit Sub
        End If
 On Error GoTo 0


    Dim oWindow, oViewer
    Set oWindow = CATIA.ActiveWindow
    Set oViewer = oWindow.ActiveViewer
    
    oWindow.Layout = catWindowGeomOnly
    
    oViewer.Reframe
    
    Dim MyViewer: Set MyViewer = CATIA.ActiveWindow.ActiveViewer
        imsg = MsgBox("是:“改为白色背景”，否：“改默认背景”，“取消退出”", vbYesNoCancel + vbDefaultButton2, "请选择操作")

        Select Case imsg
            Case 7 '===选择“否”====
                MyViewer.PutBackgroundColor Array(0.2, 0.2, 0.4)
                oWindow.Layout = catWindowSpecsAndGeom
            Case 2:
                oWindow.Layout = catWindowSpecsAndGeom
                Exit Sub '===选择“取消”退出====
                
            Case 6  '===选择“是”,改白色====
               MyViewer.PutBackgroundColor Array(1, 1, 1)
               oWindow.Layout = catWindowGeomOnly
   
        End Select
'        oWindow.Layout = catWindowSpecsAndGeom 'catWindowSpecsOnly ' catWindowGeomOnly
  

''====修改背景颜色=====
'    Dim MyViewer, oColor(2)
'    Dim icolor
'    icolor = Array(0.2, 0.2, 0.4)
'    Set MyViewer = catia.ActiveWindow.ActiveViewer
'    MyViewer.GetBackgroundColor oColor
'    MyViewer.PutBackgroundColor Array(1, 1, 1) ' Change background color to WHITE
'
''====修改背景颜色=====

'    MyViewer.PutBackgroundColor icolor ' Change background original
'
'    MsgBox ("已经保存图片")
'    oWindow.Layout = catWindowSpecsAndGeom 'catWindowSpecsOnly ' catWindowGeomOnly


ErrorHandler:
    If Err.Number <> 0 Then
        Select Case Err.Number
            Case 1001
            
                MsgBox "CATIA 程序错误：" & Err.Description, vbCritical
                Err.Clear
                Exit Sub
                
                Case 1002
                
        End Select
    
    End If

End Sub

