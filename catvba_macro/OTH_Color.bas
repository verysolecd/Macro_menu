Attribute VB_Name = "OTH_Color"
'Attribute VB_Name = "M60_Color"
'{GP:6}
'{Ep:CATmain}
'{Caption:������ɫ}
'{ControlTipText:�׺�ɫ�����л�}
'{BackColor: }
Const CF_BITMAP = 2
' ���°�ť���ֵĹ�������
Public Sub UpdateButtonText(ByVal btn As MSForms.CommandButton, ByVal isWhiteBackground As Boolean)
    If isWhiteBackground Then
        btn.Caption = "Ĭ�ϱ���"
    Else
        btn.Caption = "��ɫ����"
    End If
End Sub

Sub CATMain()
    On Error GoTo ErrorHandler
    If CATIA.Windows.Count < 1 Then
        MsgBox "û�д򿪵Ĵ���"
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
    ' ���ݵ�ǰ����ɫֱ���л�
    If currentColor(0) = 1 And currentColor(1) = 1 And currentColor(2) = 1 Then
        ' ��ǰ�ǰ�ɫ�������л���Ĭ�ϱ���
        MyViewer.PutBackgroundColor Array(0.2, 0.2, 0.4)
        oWindow.Layout = catWindowSpecsAndGeom
    Else
        ' ��ǰ��Ĭ�ϱ������л�����ɫ����
        MyViewer.PutBackgroundColor Array(1, 1, 1)
        oWindow.Layout = catWindowGeomOnly
    End If
    
    On Error GoTo 0
    
ErrorHandler:
    If Err.Number <> 0 Then
        Select Case Err.Number
            Case 1001
                MsgBox "CATIA ����" & Err.Description, vbCritical
                Err.Clear
                Exit Sub
            Case 1002
        End Select
    End If
End Sub
