Attribute VB_Name = "OTH_3Dmark"

' 为产品创建3D标识
'{GP:6}
'{EP:CATMain}
'{Caption:创建零件label}
'{ControlTipText: 点击后一次性创建零件3Dmakrtext}
'{背景颜色: 12648447}
' Purpose: Create a label on a product.

Private rPrd
Sub CATMain()
    If Not CanExecute("ProductDocument") Then Exit Sub
    Set rPrd = CATIA.ActiveDocument.Product
    Set g_allPN = KCL.InitDic
    g_allPN.RemoveAll
    recurthisPrd rPrd
End Sub

Sub recurthisPrd(oPrd)
        If g_allPN.Exists(oPrd.partNumber) = False Then
            g_allPN(oPrd.partNumber) = 1
            Call recurexcute(oPrd)
            End If
        If oPrd.Products.count > 0 Then
                For Each Product In oPrd.Products
                    Call recurthisPrd(Product)
                 Next
        End If
End Sub
Sub recurexcute(oPrd)
    Call c3Dmark(oPrd)
End Sub
Sub c3Dmark(oPrd)

If oPrd.Products.count < 1 Then
    If pdm Is Nothing Then Set pdm = New Cls_PDM
     info = pdm.infoPrd(oPrd)
        On Error GoTo 0
        Dim pos(11), sTextString, cMarker3Ds, oMarker3D
        oPrd.Position.GetComponents pos
        sTextString = info(3) & vbNewLine & _
                        info(5) & vbNewLine & _
                        info(7)
        Set cMarker3Ds = rPrd.GetTechnologicalObject("Marker3Ds")

        Dim pos1(2), pos2(2)
        pos1(0) = pos(9)
        pos1(1) = pos(10)
        pos1(2) = pos(11)
        
        pos2(0) = pos(0) - 500
        pos2(1) = pos(1) + 200
        pos2(2) = pos(2) + 500
        Set oMarker3D = cMarker3Ds.Add3DText(pos2, sTextString, pos1, oPrd)
        oMarker3D.TextSize = 6#
        oMarker3D.Update
    End If
End Sub

Sub Pt_annotation()

Set oDoc = CATIA.ActiveDocument
 Set oPrd = CATIA.ActiveDocument.Product
    Set oPrt = oDoc.part
 Set oHb = KCL.SelectItem("请选择geoset", "HybridBody")
  Set opt = oHb.HybridShapes.item(1)
Set anSets = oPrt.AnnotationSets
Set anset = anSets.Add("ISO_3D")
Set ref = oPrt.CreateReferenceFromObject(opt)

Set usfs = oPrt.UserSurfaces
Set usf = usfs.Generate(ref)
Set AnttF = anset.AnnotationFactory

Set AnttF = anset.AnnotationFactory2

'Set anote = AnttF.CreateEvoluateText(usf, 94.142136, 14.142136, 0#, True)
'anote.Text.Text = "tetx1"
oPrt.Update
' Set anote = AnttF.CreateText(usf)
 
' Set anote = AnttF.CreateTextNOA(usf)
  Set anote = AnttF.CreateFlagNote(usf)
    
   anote.Name = "an1"
'   anote.Text = "tetx1"
    anote.FlagNote = "tetx2"
End Sub

