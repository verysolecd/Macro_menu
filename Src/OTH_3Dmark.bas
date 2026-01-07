Attribute VB_Name = "OTH_3Dmark"
' 为产品创建3D标识
'{GP:6}
'{EP:CATMain}
'{Caption:创建零件label}
'{ControlTipText: 点击后一次性创建零件3Dmakrtext}
'{背景颜色: 12648447}
' Purpose: Create a label on a product.

Private rootprd
Sub CATMain()
    If Not CanExecute("ProductDocument") Then Exit Sub
    Set rootprd = CATIA.ActiveDocument.Product
    Set allPN = KCL.InitDic
    allPN.RemoveAll
    recurthisPrd rootprd
End Sub

Sub recurthisPrd(oprd)
        If allPN.Exists(oprd.PartNumber) = False Then
            allPN(oprd.PartNumber) = 1
            Call recurexcute(oprd)
            End If
       
        If oprd.Products.count > 0 Then
                For Each Product In oprd.Products
                    Call recurthisPrd(Product)
                 Next
        End If
End Sub
Sub recurexcute(oprd)
    Call c3Dmark(oprd)
End Sub
Sub c3Dmark(oprd)
    If pdm Is Nothing Then Set pdm = New Cls_PDM
     info = pdm.infoPrd(oprd)
        On Error GoTo 0
        Dim pos(11), sTextString, cMarker3Ds, oMarker3D
        oprd.Position.GetComponents pos
        sTextString = info(3) & vbNewLine & _
                        info(5) & vbNewLine & _
                        info(7)
        Set cMarker3Ds = rootprd.GetTechnologicalObject("Marker3Ds")

        Dim pos1(2), pos2(2)
        pos1(0) = pos(0)
        pos1(1) = pos(1)
        pos1(2) = pos(2)
        pos2(0) = pos(0) - 500
        pos2(1) = pos(1) + 200
        pos2(2) = pos(2) + 500
        Set oMarker3D = cMarker3Ds.Add3DText(pos2, sTextString, pos1, oprd)
        oMarker3D.TextSize = 6#
        oMarker3D.Update
End Sub

