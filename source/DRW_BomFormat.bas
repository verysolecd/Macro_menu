Attribute VB_Name = "DRW_BomFormat"
'{GP:5}
'{EP:CATMain}
'{Caption:设定BOM格式}
'{ControlTipText: 按初始化模板设定BOM格式}
'{背景颜色: 12648447}

Option Explicit

Sub CATMain()
 If Not CanExecute("ProductDocument") Then Exit Sub
Dim rootPrd: Set rootPrd = CATIA.ActiveDocument.Product
Dim Asm: Set Asm = rootPrd.getItem("BillOfMaterial")
Dim Ary(7) 'change number if you have more custom columns/array...
Ary(0) = "Number"
Ary(1) = "Part Number"
Ary(2) = "Quantity"
Ary(3) = "Nomenclature"
Ary(4) = "Defintion"
Ary(5) = "Mass"
Ary(6) = "Density"
Ary(7) = "Material"
Asm.SetCurrentFormat Ary

End Sub
