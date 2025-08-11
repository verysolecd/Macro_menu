Sub CATMain()
    Set oDoc = CATIA.ActiveDocument
    Set oPart = oDoc.part
    Set HSF = oPart.HybridShapeFactory
    '======= 要求选择body
    Dim imsg, filter(0)
    imsg = "请选择body"
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
            MsgBox "完成：" & i & "个点", vbInformation
    End If

End Sub
