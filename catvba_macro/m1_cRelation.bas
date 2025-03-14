Attribute VB_Name = "m1_cRelation"
Sub CATMain()

Dim partDocument1 As PartDocument
Set partDocument1 = CATIA.Activedocument

Dim selection1 As Selection
Set selection1 = partDocument1.Selection

selection1.Clear

Dim knowledgeObject1 As KnowledgeObject
' No resolution found for the object knowledgeObject1...

Dim anyObject1 As AnyObject
Set anyObject1 = knowledgeObject1.GetItem(">Part Info")

selection1.Add anyObject1

selection1.Copy

End Sub

