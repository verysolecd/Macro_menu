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
    Set oDoc = CATIA.ActiveDocument
    Set oprt = oDoc.part
    Set HSF = oprt.HybridShapeFactory
'======= 选择要识别的面
    Dim imsg: imsg = "选择要识别的面"
    Dim filter(0): filter(0) = "Face"
    Dim iSel: Set iSel = Nothing
    On Error Resume Next
      
    Set iSel = KCL.SelectItem(imsg, filter)
    On Error GoTo 0
    If Not iSel Is Nothing Then
        Set oHB = oprt.HybridBodies.Add()
        oHB.Name = "extracted points"
        Set oExtact = HSF.AddNewExtract(iSel)
        oHB.AppendHybridShape oExtact
        oprt.Update
        Set oref = oprt.CreateReferenceFromObject(oExtact)
        Set oFace = HSF.AddNewSurfaceDatum(oref)
        HSF.DeleteObjectForDatum oref
        Dim oBdry As HybridShapeBoundary
        Set oBdry = HSF.AddNewBoundaryOfSurface(oFace)
        oHB.AppendHybridShape oBdry
        oprt.Update
        Dim osel
        Set osel = CATIA.ActiveDocument.Selection
        osel.Clear
        osel.Add oBdry
        CATIA.StartCommand ("Disassemble")
        CATIA.RefreshDisplay = True
        MsgBox "请拆解窗口选择only domain后点击ok，再点击本窗口的ok"
          CATIA.RefreshDisplay = True
        osel.Clear
        i = 1
        For Each Hole In oHB.HybridShapes
            osel.Add Hole
            If TypeOf Hole Is HybridShapeCircleTritangent Then
                Set oref = oprt.CreateReferenceFromObject(Hole)
                Set oCtr = HSF.AddNewPointCenter(oref)
                oHB.AppendHybridShape oCtr
                Set oref = oprt.CreateReferenceFromObject(oCtr)
                oprt.Update
                Set pt = HSF.AddNewPointDatum(oref): pt.Name = "pt_" & i
                oHB.AppendHybridShape pt
                osel.Add oCtr
                i = i + 1
              Else
                osel.Add Hole
            End If
        Next
                On Error Resume Next
                    osel.Delete
                     osel.Clear
                 On Error GoTo 0
     End If
End Sub


