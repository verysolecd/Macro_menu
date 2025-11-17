Attribute VB_Name = "MDL_holecenter2"
'Attribute VB_Name = "m26_holecenter2"
' 获得识别特征下的所有孔中心
'{GP:4}
'{EP:Faceholecenter}
'{Caption:孔中心点}
'{ControlTipText: 提示选择面后后导出面上所有孔中心}
'{BackColor:12648447}

Sub Faceholecenter()
    If CATIA.Windows.count < 1 Then
        MsgBox "没有打开的窗口"
        Exit Sub
    End If
  If Not CanExecute("PartDocument") Then Exit Sub
    Set odoc = CATIA.ActiveDocument
    Set oPart = odoc.part
    Set HSF = oPart.HybridShapeFactory
'======= 选择要识别的面
    Dim imsg, filter(0), iSel
    imsg = "选择要识别的面"
    filter(0) = "Face"
    On Error Resume Next
    Set iSel = Nothing
    Set iSel = KCL.SelectElement(imsg, filter).value
    On Error GoTo 0
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
        Dim oSel
        Set oSel = CATIA.ActiveDocument.Selection
        oSel.Clear
        oSel.Add oBdry
        CATIA.StartCommand ("Disassemble")
        CATIA.RefreshDisplay = True
        MsgBox "请拆解窗口选择only domain后点击ok，再点击本窗口的ok"
          CATIA.RefreshDisplay = True
        oSel.Clear
        i = 1
        For Each Hole In oHB.HybridShapes
            oSel.Add Hole
            If TypeOf Hole Is HybridShapeCircleTritangent Then
                Set oref = oPart.CreateReferenceFromObject(Hole)
                Set oCtr = HSF.AddNewPointCenter(oref)
                oHB.AppendHybridShape oCtr
                Set oref = oPart.CreateReferenceFromObject(oCtr)
                oPart.Update
                Set pt = HSF.AddNewPointDatum(oref)
                pt.Name = "pt_" & i
                oHB.AppendHybridShape pt
                'oPart.Update
                oSel.Add oCtr
'                osel.Delete
'                osel.Clear
                i = i + 1
              Else
                oSel.Add Hole
                
            End If
            
        Next
                On Error Resume Next
                oSel.Delete
                 oSel.Clear
                 On Error GoTo 0
     End If
End Sub


