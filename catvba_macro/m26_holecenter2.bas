Attribute VB_Name = "m26_holecenter2"
'Attribute VB_Name = "m26_holecenter2"
' 获得识别特征下的所有孔中心
'{GP:2}
'{EP:Faceholecenter}
'{Caption:孔中心点}
'{ControlTipText: 提示选择面后后导出面上所有孔中心}
'{BackColor:12648447}

Sub Faceholecenter()
    If CATIA.Windows.Count < 1 Then
        MsgBox "没有打开的窗口"
        Exit Sub
    End If
  If Not CanExecute("PartDocument") Then Exit Sub
    Set oDoc = CATIA.ActiveDocument
    Set oPart = oDoc.part
    Set HSF = oPart.HybridShapeFactory
'======= 选择要识别的面
    Dim imsg, filter(0), iSel
    imsg = "选择要识别的面"
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
        MsgBox "请拆解窗口选择only domain后点击ok，再点击本窗口的ok"
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


