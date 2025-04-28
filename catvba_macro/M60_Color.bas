Attribute VB_Name = "M60_Color"
'{GP:6}
'{Ep:CATmain}
'{Caption:��ɫ����}
'{ControlTipText:�İ�ɫ�����ý�ͼ}
'{BackColor: }

Const CF_BITMAP = 2
Sub CATMain()
 On Error GoTo ErrorHandler
        If CATIA.Documents.Count = 0 Then
            Err.Raise 1001, , "δ��⵽�򿪵�CATIA�ĵ�"
            Exit Sub
        End If
 On Error GoTo 0


    Dim oWindow, oViewer
    Set oWindow = CATIA.ActiveWindow
    Set oViewer = oWindow.ActiveViewer
    
    oWindow.Layout = catWindowGeomOnly
    
    oViewer.Reframe
    
    Dim MyViewer: Set MyViewer = CATIA.ActiveWindow.ActiveViewer
        imsg = MsgBox("��:����Ϊ��ɫ���������񣺡���Ĭ�ϱ���������ȡ���˳���", vbYesNoCancel + vbDefaultButton2, "��ѡ�����")

        Select Case imsg
            Case 7 '===ѡ�񡰷�====
                MyViewer.PutBackgroundColor Array(0.2, 0.2, 0.4)
                oWindow.Layout = catWindowSpecsAndGeom
            Case 2:
                oWindow.Layout = catWindowSpecsAndGeom
                Exit Sub '===ѡ��ȡ�����˳�====
                
            Case 6  '===ѡ���ǡ�,�İ�ɫ====
               MyViewer.PutBackgroundColor Array(1, 1, 1)
               oWindow.Layout = catWindowGeomOnly
   
        End Select
'        oWindow.Layout = catWindowSpecsAndGeom 'catWindowSpecsOnly ' catWindowGeomOnly
  

''====�޸ı�����ɫ=====
'    Dim MyViewer, oColor(2)
'    Dim icolor
'    icolor = Array(0.2, 0.2, 0.4)
'    Set MyViewer = catia.ActiveWindow.ActiveViewer
'    MyViewer.GetBackgroundColor oColor
'    MyViewer.PutBackgroundColor Array(1, 1, 1) ' Change background color to WHITE
'
''====�޸ı�����ɫ=====

'    MyViewer.PutBackgroundColor icolor ' Change background original
'
'    MsgBox ("�Ѿ�����ͼƬ")
'    oWindow.Layout = catWindowSpecsAndGeom 'catWindowSpecsOnly ' catWindowGeomOnly


ErrorHandler:
    If Err.Number <> 0 Then
        Select Case Err.Number
            Case 1001
            
                MsgBox "CATIA �������" & Err.Description, vbCritical
                Err.Clear
                Exit Sub
                
                Case 1002
                
        End Select
    
    End If

End Sub

