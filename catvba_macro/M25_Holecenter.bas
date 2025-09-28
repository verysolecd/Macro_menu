Attribute VB_Name = "M25_Holecenter"
'Attribute VB_Name = "M25_Holecenter"
' ���ʶ�������µ����п�����
'{GP:2}
'{EP:ctrhole}
'{Caption:get�����ĵ�}
'{ControlTipText: ��ʾѡ��ʵ��󵼳����п����ģ�������ʶ����������ʵ��}
'{BackColor:12648447}

Sub ctrhole()

 If CATIA.Windows.Count < 1 Then
        MsgBox "û�д򿪵Ĵ���"
        Exit Sub
    End If
    
  If Not CanExecute("PartDocument") Then Exit Sub

    Set oDoc = CATIA.ActiveDocument
    Set oPart = oDoc.part
    Set HSF = oPart.HybridShapeFactory
    '======= Ҫ��ѡ��body
    Dim imsg, filter(0)
    imsg = "��ѡ��body"
    filter(0) = "Body"
    Dim obdy
    Set obdy = KCL.SelectElement(imsg, filter).Value
    Set targethb = oPart.HybridBodies.Add()
    targethb.Name = "extracted points"
    If Not obdy Is Nothing Then
            Set holeBody = obdy
            For Each Hole In holeBody.Shapes
            If TypeOf Hole Is Hole Then
                Set skt = Hole.Sketch
                Set Pt = HSF.AddNewPointCoord(0, 0, 0)
                Set ref = oPart.CreateReferenceFromObject(skt)
                Pt.PtRef = ref
                Pt.Name = "Pt_" & i
                targethb.AppendHybridShape Pt
                oPart.InWorkObject = Pt
                oPart.Update
                i = i + 1
            End If
        Next
        MsgBox "��ɣ�" & i & "����", vbInformation
    End If
End Sub
