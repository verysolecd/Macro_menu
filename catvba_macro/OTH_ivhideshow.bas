Attribute VB_Name = "OTH_ivhideshow"
'Attribute VB_Name = "m64_hide&show"
'{GP:6}
'{Ep:CATMain}
'{Caption:反选隐藏}
'{ControlTipText:反选并隐藏结构树}
'{BackColor:}

Sub CATMain()
Dim oDoc
Set oDoc = CATIA.ActiveDocument
Dim cGroups
Set cGroups = oDoc.Product.GetTechnologicalObject("Groups")
Dim oGroup As Group
Set oGroup = cGroups.AddFromSel    ' 当前选择产品添加到组
oGroup.ExtractMode = 1
oGroup.FillSelWithInvert   '  反选
'oGroup.FillSelWithExtract
' Delete the group
  cGroups.Remove 1
  Set cGroups = Nothing
  Dim sel
Set sel = oDoc.Selection
Set VisPropertySet = sel.VisProperties
sel.VisProperties.SetShow 1  '' 将所有选中元素设置为不可见
'VisPropertySet.SetShow 0
End Sub
