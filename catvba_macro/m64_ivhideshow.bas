Attribute VB_Name = "m64_ivhideshow"
'Attribute VB_Name = "m64_hide&show"
'{GP:6}
'{Ep:CATMain}
'{Caption:��ѡ����}
'{ControlTipText:��ѡ�����ؽṹ��}
'{BackColor:}

Sub CATMain()
Dim odoc
Set odoc = CATIA.ActiveDocument
Dim cGroups
Set cGroups = odoc.product.GetTechnologicalObject("Groups")
Dim oGroup As Group
Set oGroup = cGroups.AddFromSel    ' ��ǰѡ���Ʒ��ӵ���
oGroup.ExtractMode = 1
oGroup.FillSelWithInvert   '  ��ѡ
'oGroup.FillSelWithExtract
' Delete the group
  cGroups.Remove 1
  Set cGroups = Nothing
  Dim sel
Set sel = odoc.Selection
Set VisPropertySet = sel.VisProperties
sel.VisProperties.SetShow 1  '' ������ѡ��Ԫ������Ϊ���ɼ�
'VisPropertySet.SetShow 0
End Sub
