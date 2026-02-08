Attribute VB_Name = "OTH_3Dmark"


' 为产品创建3D标识
'{GP:6}
'{EP:newlabel}
'{Caption:创建零件label}
'{ControlTipText: 点击后一次性创建零件3Dmakrtext}
'{背景颜色: 12648447}
' Purpose: Create a label on a product.

Private rprd
Private Const mdlname As String = "OTH_3Dmark"
Sub newlabel()
    If Not CanExecute("ProductDocument") Then Exit Sub
    Set rprd = CATIA.ActiveDocument.Product
    Set g_allPN = KCL.InitDic
    g_allPN.RemoveAll
    recurthisPrd rprd
End Sub

Sub recurthisPrd(oPrd)
        If g_allPN.Exists(oPrd.PartNumber) = False Then
            g_allPN(oPrd.PartNumber) = 1
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
Dim info As Bomline
   info = pdm.getBomLine(oPrd)
        On Error GoTo 0
        Dim pos(11), sTextString, cMarker3Ds, oMarker3D
        oPrd.Position.GetComponents pos
        
        ' Updated for Bomline Typed return
        Dim def As String, mat As String, massStr As String
        def = info.Definition
        mat = info.Material
        massStr = Format(info.Mass, "0.000") & " kg" ' Format mass nicely
        
        sTextString = def & vbNewLine & _
                        mat & vbNewLine & _
                        massStr
        Set cMarker3Ds = rprd.GetTechnologicalObject("Marker3Ds")

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
    Set oprt = oDoc.part
 Set oHb = KCL.SelectItem("请选择geoset", "HybridBody")
  Set opt = oHb.HybridShapes.item(1)
Set anSets = oprt.AnnotationSets
Set anset = anSets.Add("ISO_3D")
Set ref = oprt.CreateReferenceFromObject(opt)

Set usfs = oprt.UserSurfaces
Set usf = usfs.Generate(ref)
Set AnttF = anset.AnnotationFactory

Set AnttF = anset.AnnotationFactory2

'Set anote = AnttF.CreateEvoluateText(usf, 94.142136, 14.142136, 0#, True)
'anote.Text.Text = "tetx1"
oprt.Update
' Set anote = AnttF.CreateText(usf)
 
' Set anote = AnttF.CreateTextNOA(usf)
  Set anote = AnttF.CreateFlagNote(usf)
    
   anote.name = "an1"
'   anote.Text = "tetx1"
    anote.FlagNote = "tetx2"
End Sub



