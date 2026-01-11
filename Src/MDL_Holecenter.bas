Attribute VB_Name = "MDL_Holecenter"
'Attribute VB_Name = "M25_Holecenter"
' 获得识别特征下的所有孔中心
'{GP:}
'{EP:ctrhole}
'{Caption:get孔中心点}
'{ControlTipText: 提示选择实体后导出所有孔中心，必须是识别孔特征后的实体}
'{BackColor:12648447}

Sub ctrhole()
    
  If Not CanExecute("PartDocument") Then Exit Sub

    Set oDoc = CATIA.ActiveDocument
    Set oPart = oDoc.part
    Set HSF = oPart.HybridShapeFactory
    '======= 要求选择body
    Dim imsg, filter(0)
    imsg = "请选择body"
    filter(0) = "Body"
    Dim obdy
    Set obdy = KCL.SelectItem(imsg, filter)
    Set targetHB = oPart.HybridBodies.Add()
    targetHB.Name = "extracted points"
    If Not obdy Is Nothing Then
            Set holeBody = obdy
            For Each Hole In holeBody.Shapes
                If TypeOf Hole Is Hole Then
                    Set skt = Hole.Sketch
                    Set pt = HSF.AddNewPointCoord(0, 0, 0)
                    Set ref = oPart.CreateReferenceFromObject(skt)
                    pt.PtRef = ref
                    pt.Name = "Pt_" & i
                    targetHB.AppendHybridShape pt
                    oPart.InWorkObject = pt
                    oPart.Update
                    i = i + 1
                End If
            Next
        MsgBox "完成：" & i & "个点", vbInformation
    End If
End Sub
