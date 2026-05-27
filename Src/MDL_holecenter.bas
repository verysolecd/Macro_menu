Attribute VB_Name = "MDL_holecenter"
'Attribute VB_Name = "m26_holecenter"
' 获得识别特征下的所有孔中心
'{GP:4}
'{EP:Faceholecenter}
'{Caption:获取all孔中心}
'{ControlTipText: 提示选择面后后导出面上所有孔中心}
'{BackColor: }

Private Const mdlname As String = "MDL_holecenter"
Sub Faceholecenter()
 If Not CanExecute("Productdocument,PartDocument") Then Exit Sub
    On Error Resume Next
        Dim oDoc: Set oDoc = CATIA.ActiveDocument
        Dim workPrtDoc: Set workPrtDoc = KCL.get_workPartDoc
        Dim oprt: Set oprt = Nothing: Set oprt = workPartDoc.part
    Err.Clear
    On Error GoTo 0
    If IsNothing(oprt) Then: MsgBox "No activated Part": Exit Sub
    Set HSF = oprt.HybridShapeFactory
'======= 选择要识别的面
Dim isel: Set isel = Nothing
    Dim imsg: imsg = "选择要识别的面"
    Dim filter(0): filter(0) = "Face,HybridShape"
    On Error Resume Next
        Set isel = KCL.SelectItem(imsg, filter)
    On Error GoTo 0
    If Not isel Is Nothing Then
        Set oHB = oprt.HybridBodies.Add(): oHB.Name = "extracted points"
        Set oExtact = HSF.AddNewExtract(isel)
            oHB.AppendHybridShape oExtact
            oprt.Update
        Set oRef = oprt.CreateReferenceFromObject(oExtact)
        Set oFace = HSF.AddNewSurfaceDatum(oRef)
            HSF.DeleteObjectForDatum oRef
        Dim oBdry As HybridShapeBoundary: Set oBdry = HSF.AddNewBoundaryOfSurface(oFace)
            oHB.AppendHybridShape oBdry
        oprt.Update
        Dim osel: Set osel = CATIA.ActiveDocument.Selection
        osel.Clear: osel.Add oBdry
            CATIA.StartCommand ("Disassemble")
            CATIA.RefreshDisplay = True
                MsgBox "请拆解窗口选择only domain后点击ok，再点击本窗口的ok"
            CATIA.RefreshDisplay = False
        osel.Clear
        i = 1
        For Each Hole In oHB.HybridShapes
            osel.Add Hole
            If TypeOf Hole Is HybridShapeCircleTritangent Then
                Set oRef = oprt.CreateReferenceFromObject(Hole)
                Set oCtr = HSF.AddNewPointCenter(oRef)
                oHB.AppendHybridShape oCtr
                Set oRef = oprt.CreateReferenceFromObject(oCtr)
                oprt.Update
                Set pt = HSF.AddNewPointDatum(oRef): pt.Name = "pt_" & i
                oHB.AppendHybridShape pt
                osel.Add oCtr
                i = i + 1
              Else
                osel.Add Hole
            End If
        Next
                On Error Resume Next
                    osel.Delete: osel.Clear
                 On Error GoTo 0
     End If
     
     CATIA.RefreshDisplay = True
     Set osel = Nothing
     Set isel = Nothing
End Sub


