Attribute VB_Name = "MDL_holecenter2"
'Attribute VB_Name = "m26_holecenter2"
' 获得识别特征下的所有孔中心
'{GP:4}
'{EP:Faceholecenter}
'{Caption:孔中心点}
'{ControlTipText: 提示选择面后后导出面上所有孔中心}
'{BackColor:12648447}

Private Const mdlname As String = "MDL_holecenter2"
Sub Faceholecenter()
  If Not CanExecute("PartDocument,ProductDocument") Then Exit Sub

    Set oprt = KCL.get_workPartDoc.part
    Set HSF = oprt.HybridShapeFactory
'======= 选择要识别的面
Dim iSel: Set iSel = Nothing
    Dim imsg: imsg = "选择要识别的面"
    Dim filter(0): filter(0) = "Face"
    On Error Resume Next
        Set iSel = KCL.SelectItem(imsg, filter)
    On Error GoTo 0
    If Not iSel Is Nothing Then
        Set oHb = oprt.HybridBodies.Add(): oHb.name = "extracted points"
        Set oExtact = HSF.AddNewExtract(iSel)
            oHb.AppendHybridShape oExtact
            oprt.Update
        Set oRef = oprt.CreateReferenceFromObject(oExtact)
        Set oFace = HSF.AddNewSurfaceDatum(oRef)
            HSF.DeleteObjectForDatum oRef
        Dim oBdry As HybridShapeBoundary: Set oBdry = HSF.AddNewBoundaryOfSurface(oFace)
            oHb.AppendHybridShape oBdry
        oprt.Update
        Dim oSel: Set oSel = CATIA.ActiveDocument.Selection
        oSel.Clear: oSel.Add oBdry
            CATIA.StartCommand ("Disassemble")
            CATIA.RefreshDisplay = True
                MsgBox "请拆解窗口选择only domain后点击ok，再点击本窗口的ok"
            CATIA.RefreshDisplay = False
        oSel.Clear
        i = 1
        For Each Hole In oHb.HybridShapes
            oSel.Add Hole
            If TypeOf Hole Is HybridShapeCircleTritangent Then
                Set oRef = oprt.CreateReferenceFromObject(Hole)
                Set oCtr = HSF.AddNewPointCenter(oRef)
                oHb.AppendHybridShape oCtr
                Set oRef = oprt.CreateReferenceFromObject(oCtr)
                oprt.Update
                Set pt = HSF.AddNewPointDatum(oRef): pt.name = "pt_" & i
                oHb.AppendHybridShape pt
                oSel.Add oCtr
                i = i + 1
              Else
                oSel.Add Hole
            End If
        Next
                On Error Resume Next
                    oSel.Delete: oSel.Clear
                 On Error GoTo 0
     End If
     
     CATIA.RefreshDisplay = True
     Set oSel = Nothing
     Set iSel = Nothing
End Sub


