Attribute VB_Name = "m26_holecenter2"
'Attribute VB_Name = "m26_holecenter2"
' ���ʶ�������µ����п�����
'{GP:2}
'{EP:Faceholecenter}
'{Caption:�����ĵ�}
'{ControlTipText: ��ʾѡ�����󵼳��������п�����}
'{BackColor:12648447}

Sub Faceholecenter()
    If CATIA.Windows.Count < 1 Then
        MsgBox "û�д򿪵Ĵ���"
        Exit Sub
    End If
  If Not CanExecute("PartDocument") Then Exit Sub
    Set oDoc = CATIA.ActiveDocument
    Set oPart = oDoc.part
    Set HSF = oPart.HybridShapeFactory
'======= ѡ��Ҫʶ�����
    Dim imsg, filter(0), iSel
    imsg = "ѡ��Ҫʶ�����"
    filter(0) = "Face"
    Set iSel = KCL.SelectElement(imsg, filter).Value
    
    If Not iSel Is Nothing Then
        Set oHB = oPart.HybridBodies.Add()
        oHB.Name = "extracted points"
        Set oExtact = HSF.AddNewExtract(iSel)
        oHB.AppendHybridShape oExtact
        oPart.Update
        Set oref = oPart.CreateReferenceFromObject(oExtact)
        Set oFace = HSF.AddNewSurfaceDatum(oref)
        HSF.DeleteObjectForDatum oref
        Dim oBdry As HybridShapeBoundary
        Set oBdry = HSF.AddNewBoundaryOfSurface(oFace)
        oHB.AppendHybridShape oBdry
        oPart.Update
        Dim osel
        Set osel = CATIA.ActiveDocument.Selection
        osel.Clear
        osel.Add oBdry
        CATIA.StartCommand ("Disassemble")
        CATIA.RefreshDisplay = True
        MsgBox "���ⴰ��ѡ��only domain����ok���ٵ�������ڵ�ok"
          CATIA.RefreshDisplay = True
        osel.Clear
        For Each Hole In oHB.HybridShapes
            osel.Add Hole
            If TypeOf Hole Is HybridShapeCircleTritangent Then
                Set oref = oPart.CreateReferenceFromObject(Hole)
                Set oCtr = HSF.AddNewPointCenter(oref)
                oHB.AppendHybridShape oCtr
                Set oref = oPart.CreateReferenceFromObject(oCtr)
                oPart.Update
                Set pt = HSF.AddNewPointDatum(oref)
                oHB.AppendHybridShape pt
                oPart.Update
                osel.Add oCtr
'                osel.Delete
'                osel.Clear
              Else
                osel.Add Hole
                
            End If
        Next
        osel.Delete
                 osel.Clear
     End If
End Sub


