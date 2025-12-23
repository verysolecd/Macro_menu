Attribute VB_Name = "Module1"
Sub CATMain()

Dim documents1 As Documents
Set documents1 = CATIA.Documents

Dim productDocument1 As ProductDocument
Set productDocument1 = documents1.item("_1100.CATProduct")

Dim product1 As Product
Set product1 = productDocument1.Product

Set product1 = product1.ReferenceProduct

Dim productDocument2 As ProductDocument
Set productDocument2 = documents1.item("_1000.CATProduct")

Dim product2 As Product
Set product2 = productDocument2.Product

Dim products1 As Products
Set products1 = product2.Products

Dim product3 As Product
Set product3 = products1.item("Product1.1")

product3.Name = "Ãû×Ö"

End Sub

