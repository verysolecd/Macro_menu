Attribute VB_Name = "getsize1"
Sub getSize()
'If Not KCL.CanExecute("PartDocument") Then Exit Sub

CATIA.HSOSynchronized = False

Dim doc: Set doc = CATIA.ActiveDocument
Dim oprt
Set oprt = doc.Product.Products.item(1).ReferenceProduct.Parent.part

Dim MBD
Set MBD = oprt.bodies.item(1)

Dim BBX 'As HybridBody
Set BBX = oprt.HybridBodies.Add
BBX.Name = "Boundingbox"

Dim ref
Set ref = oprt.CreateReferenceFromObject(MBD)
Dim HSF 'As HybridShapeFactory
Set HSF = oprt.HybridShapeFactory

Dim Extract
Set Extract = HSF.AddNewExtract(ref)
BBX.AppendHybridShape Extract

Extract.Compute
doc.Selection.Add Extract

Set ref = oprt.CreateReferenceFromObject(Extract)
Dim xDir
Set xDir = HSF.AddNewDirectionByCoord(1, 0, 0)
Dim yDir
Set yDir = HSF.AddNewDirectionByCoord(0, 1, 0)
Dim zDir
Set zDir = HSF.AddNewDirectionByCoord(0, 0, 1)

Dim xmax, xmin, ymax, ymin, zmax, zmin

Set xmax = HSF.AddNewExtremum(ref, xDir, 1)
BBX.AppendHybridShape xmax
doc.Selection.Add xmax

Set xmin = HSF.AddNewExtremum(ref, xDir, 0)
BBX.AppendHybridShape xmin
doc.Selection.Add xmin

Set ymax = HSF.AddNewExtremum(ref, yDir, 1)
BBX.AppendHybridShape ymax
doc.Selection.Add ymax

Set ymin = HSF.AddNewExtremum(ref, yDir, 0)
BBX.AppendHybridShape ymin
doc.Selection.Add ymin

Set zmax = HSF.AddNewExtremum(ref, zDir, 1)
BBX.AppendHybridShape zmax
doc.Selection.Add zmax

Set zmin = HSF.AddNewExtremum(ref, zDir, 0)
BBX.AppendHybridShape zmin
doc.Selection.Add zmin

oprt.Update

Dim WB
Set WB = doc.GetWorkbench("SPAWorkbench")
Dim Mes(2), Arr(5), DisX, DisY, DisZ
Set Mes(0) = WB.GetMeasurable(xmax)
Mes(0).GetMinimumDistancePoints xmin, Arr
DisX = Abs(Arr(3) - Arr(0))
xmaxc = Arr(0): xminc = Arr(3)

Set Mes(1) = WB.GetMeasurable(ymax)
Mes(1).GetMinimumDistancePoints ymin, Arr
DisY = Abs(Arr(4) - Arr(1))
ymaxc = Arr(1): yminc = Arr(4)

Set Mes(2) = WB.GetMeasurable(zmax)
Mes(2).GetMinimumDistancePoints zmin, Arr
DisZ = Abs(Arr(5) - Arr(2))
zmaxc = Arr(2): zminc = Arr(5)

'    Doc.Selection.Add BBX
'    Doc.Selection.Delete

Set product2 = oprt 'Doc.part
Set parameters1 = product2.Parameters 'UserRefProperties
Dim length1 As Length, length2 As Length, length3 As Length
Set length1 = parameters1.CreateDimension("X向", "LENGTH", DisX)
Set length2 = parameters1.CreateDimension("Y向", "LENGTH", DisY)
Set length3 = parameters1.CreateDimension("Z向", "LENGTH", DisZ)
Set p1 = HSF.AddNewPointCoord(xmaxc, yminc, zminc)
BBX.AppendHybridShape p1
p1.Compute
doc.Selection.Add p1
Set p2 = HSF.AddNewPointCoord(xminc, yminc, zminc)
BBX.AppendHybridShape p2
p2.Compute
doc.Selection.Add p2
Set ln = HSF.AddNewLinePtPt(p1, p2)
BBX.AppendHybridShape ln
ln.Compute
doc.Selection.Add ln
Set ext = HSF.AddNewExtrude(ln, DisY, 0, yDir)
BBX.AppendHybridShape ext
ext.Compute
doc.Selection.Add ext
Set bound = HSF.AddNewBoundaryOfSurface(ext)
BBX.AppendHybridShape bound
bound.Compute
doc.Selection.Add bound
Set ext2 = HSF.AddNewExtrude(bound, DisZ, 0, zDir)
BBX.AppendHybridShape ext2
ext2.Compute
doc.Selection.Add ext2
Set trans = HSF.AddNewTranslate(ext, zDir, DisZ)
BBX.AppendHybridShape trans
trans.Compute
doc.Selection.Add trans
Set asm = HSF.AddNewJoin(ext, ext2)
asm.AddElement trans
BBX.AppendHybridShape asm
asm.Compute
eles = HSF.AddNewDatums(asm)
BBX.AppendHybridShape eles(0)
eles(0).Name = "Bounding box of " & MBD.Name
HSF.DeleteObjectForDatum asm

doc.Selection.Delete
doc.Selection.Add eles(0)
doc.Selection.VisProperties.SetRealOpacity 100, 1
doc.Selection.Clear
CATIA.HSOSynchronized = True

'resp = MsgBox("X direction size is " & Round(DisX, 3) & vbCrLf & _
'        "Y direction size is " & Round(DisY, 3) & vbCrLf & _
'        "Z direction size is " & Round(DisZ, 3) & vbCrLf & _
'        "Do you want to keep the bounding box geometry?", vbYesNo)
'If resp = vbNo Then
'    Doc.Selection.Add BBX
'    Doc.Selection.Delete
'Else
'    Doc.Selection.Clear
'End If
'MsgBox "零件长宽高属性已添加"
End Sub

