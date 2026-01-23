Attribute VB_Name = "MDL_wfrename"
'Attribute VB_Name = "m24_wfRename"
' 线框元素的重命名
'{GP:4}
'{EP:wfname}
'{Caption:线框重命名}
'{ControlTipText: 提示选择几何图形集后将下面元素重命名}
'{BackColor:12648447}
'type definition
' = 0 , Unknown
' = 1 , Point
' = 2 , Curve
' = 3 , Line
' = 4 , Circle
' = 5 , Surface

Private Const mdlname As String = "MDL_wfrename"
Sub wfname()
If CATIA.Windows.count < 1 Then
MsgBox "没有打开的窗口"
Exit Sub
End If
Dim oDoc
On Error Resume Next
Set oDoc = CATIA.ActiveDocument
On Error GoTo 0
Dim str
str = TypeName(oDoc)
If Not str = "PartDocument" Then
MsgBox "没有打开的part"
Exit Sub
End If
Dim HSF:  Set HSF = oDoc.part.HybridShapeFactory
Dim HBS: Set HBS = oDoc.part.HybridBodies
Dim osel: Set osel = oDoc.Selection
osel.Clear
'=======要求选择几何图形集和坐标
Dim imsg
imsg = "请选择元素所在的几何图形集"
Dim oHb
Dim filter(0)
filter(0) = "HybridBody"
Set oHb = KCL.SelectItem(imsg, filter)
If Not oHb Is Nothing Then
Dim i, qty
Set oshapes = oHb.HybridShapes
qty = oshapes.count
Dim ct  As Variant
ct = Array(0, 0, 0, 0, 0, 0, 0, 0)
Dim oWF
For i = 1 To qty
Set oWF = oshapes.item(i)
str = HSF.GetGeometricalFeatureType(oWF)
Select Case str
Case 0
    oWF.Name = "aShape" & ct(0)
    ct(0) = ct(0) + 1
Case 1
     oWF.Name = "point" & ct(1)
    ct(1) = ct(1) + 1
Case 2
   oWF.Name = "curve" & ct(2)
    ct(2) = ct(2) + 1
Case 3
  oWF.Name = "line" & ct(3)
    ct(3) = ct(3) + 1
Case 4
   oWF.Name = "circle" & ct(4)
    ct(4) = ct(4) + 1
Case 5
   oWF.Name = "surface" & ct(5)
    ct(5) = ct(5) + 1
Case 6
   oWF.Name = "plane" & ct(6)
   ct(6) = ct(6) + 1
Case 7
      oWF.Name = "solid" & ct(7)
   ct(7) = ct(7) + 1
End Select

Next
End If
End Sub
