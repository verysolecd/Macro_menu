Attribute VB_Name = "MDL_wfrename"
'Attribute VB_Name = "m24_wfrename"
' ����������Ԫ��
'{GP:4}
'{EP:wfrename}
'{Caption:����������}
'{ControlTipText: ��ʾѡ�񼸺�ͼ�μ��󵼳�����ĵ㼯}
'{BackColor:12648447}
' = 0 , Unknown
' = 1 , Point
' = 2 , Curve
' = 3 , Line
' = 4 , Circle
' = 5 , Surface
' = 6 , Plane
' = 7 , Solid, Volume

Sub wfrename()
   
  If CATIA.Windows.Count < 1 Then
        MsgBox "û�д򿪵Ĵ���"
        Exit Sub
    End If
    
    Dim oDoc
    On Error Resume Next
        Set oDoc = CATIA.ActiveDocument
    On Error GoTo 0
    Dim Str
    Str = TypeName(oDoc)
    If Not Str = "PartDocument" Then
    MsgBox "û�д򿪵�part"
    Exit Sub
    End If
 

End Sub
