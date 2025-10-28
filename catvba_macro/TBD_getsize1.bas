Attribute VB_Name = "TBD_getsize1"
Sub getSize()
'If Not KCL.CanExecute("PartDocument") Then Exit Sub

CATIA.HSOSynchronized = False

Set doc = CATIA.ActiveDocument
Set oprt = doc.part 'Product.Products.item(1).ReferenceProduct.Parent.part
Set MBD = oprt.bodies.item(1)
Set HB = oprt.HybridBodies.Add
HB.Name = "Boundingbox"
Set ref = oprt.CreateReferenceFromObject(MBD)
Dim HSF 'As HybridShapeFactory
Set HSF = oprt.HybridShapeFactory
Set extract = HSF.AddNewExtract(ref)
HB.AppendHybridShape extract
extract.Compute
doc.Selection.Add extract

Set ref = oprt.CreateReferenceFromObject(extract)

Set xDir = HSF.AddNewDirectionByCoord(1, 0, 0)
Set yDir = HSF.AddNewDirectionByCoord(0, 1, 0)
Set zDir = HSF.AddNewDirectionByCoord(0, 0, 1)

Set xmax = HSF.AddNewExtremum(ref, xDir, 1)

HB.AppendHybridShape xmax

doc.Selection.Add xmax

Set xmin = HSF.AddNewExtremum(ref, xDir, 0)

HB.AppendHybridShape xmin
doc.Selection.Add xmin

Set ymax = HSF.AddNewExtremum(ref, yDir, 1)

HB.AppendHybridShape ymax
doc.Selection.Add ymax

Set ymin = HSF.AddNewExtremum(ref, yDir, 0)

HB.AppendHybridShape ymin
doc.Selection.Add ymin

Set zmax = HSF.AddNewExtremum(ref, zDir, 1)

HB.AppendHybridShape zmax
doc.Selection.Add zmax

Set zmin = HSF.AddNewExtremum(ref, zDir, 0)

HB.AppendHybridShape zmin
doc.Selection.Add zmin

oprt.Update

Set WB = doc.GetWorkbench("SPAWorkbench")
Dim Mes(2), Arr(5), DisX, DisY, DisZ

Set Mes(0) = WB.GetMeasurable(xmax)

Mes(0).GetMinimumDistancePoints xmin, Arr
DisX = Abs(Arr(3) - Arr(0))

xmaxc = Arr(0)
xminc = Arr(3)

Set Mes(1) = WB.GetMeasurable(ymax)
Mes(1).GetMinimumDistancePoints ymin, Arr
DisY = Abs(Arr(4) - Arr(1))
ymaxc = Arr(1): yminc = Arr(4)

Set Mes(2) = WB.GetMeasurable(zmax)
Mes(2).GetMinimumDistancePoints zmin, Arr
DisZ = Abs(Arr(5) - Arr(2))
zmaxc = Arr(2): zminc = Arr(5)

'    Doc.Selection.Add HB
'    Doc.Selection.Delete

Set product2 = oprt 'Doc.part
Set parameters1 = product2.Parameters 'UserRefProperties
Dim length1 As Length, length2 As Length, length3 As Length
Set length1 = parameters1.CreateDimension("X��", "LENGTH", DisX)
Set length2 = parameters1.CreateDimension("Y��", "LENGTH", DisY)
Set length3 = parameters1.CreateDimension("Z��", "LENGTH", DisZ)

Set p1 = HSF.AddNewPointCoord(xmaxc, yminc, zminc)
HB.AppendHybridShape p1
p1.Compute
doc.Selection.Add p1

Set p2 = HSF.AddNewPointCoord(xminc, yminc, zminc)
HB.AppendHybridShape p2
p2.Compute
doc.Selection.Add p2

Set ln = HSF.AddNewLinePtPt(p1, p2)
HB.AppendHybridShape ln
ln.Compute

doc.Selection.Add ln
Set ext = HSF.AddNewExtrude(ln, DisY, 0, yDir)
HB.AppendHybridShape ext
ext.Compute
doc.Selection.Add ext
Set bound = HSF.AddNewBoundaryOfSurface(ext)
HB.AppendHybridShape bound
bound.Compute
doc.Selection.Add bound
Set ext2 = HSF.AddNewExtrude(bound, DisZ, 0, zDir)
HB.AppendHybridShape ext2
ext2.Compute
doc.Selection.Add ext2
Set trans = HSF.AddNewTranslate(ext, zDir, DisZ)
HB.AppendHybridShape trans
trans.Compute
doc.Selection.Add trans
Set asm = HSF.AddNewJoin(ext, ext2)
asm.AddElement trans
HB.AppendHybridShape asm
asm.Compute
eles = HSF.AddNewDatums(asm)
HB.AppendHybridShape eles(0)
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
'    Doc.Selection.Add HB
'    Doc.Selection.Delete
'Else
'    Doc.Selection.Clear
'End If
'MsgBox "�����������������"
End Sub

