Attribute VB_Name = "A0TEST"
Sub CATIASearchExample()
 Set oprt = CATIA.ActiveDocument.part
Set paras = oprt.Parameters
Set strParam1 = paras.item("String.3")
With strParam1
Select Case oprt.Parent.Product.ReferenceProduct.Source

Case 0: .Value = "unKnown"
    
Case 1: .Value = "Make"
Case 2: .Value = "Buy"

 
End Sub
