Attribute VB_Name = "DRW_BomFormat"
'{GP:5}
'{EP:CATMain}
'{Caption:设定BOM格式}
'{ControlTipText: 按初始化模板设定BOM格式}
'{背景颜色: 12648447}

Option Explicit

Sub CATMain()
 If Not CanExecute("ProductDocument") Then Exit Sub
    Dim rootprd: Set rootprd = CATIA.ActiveDocument.Product
    Dim Asm: Set Asm = rootprd.getItem("BillOfMaterial")
    Dim ary(7) 'change number if you have more custom columns/array...
    ary(0) = "Number"
    ary(1) = "Part Number"
    ary(2) = "Quantity"
    ary(3) = "Nomenclature"
    ary(4) = "Defintion"
    ary(5) = "Mass"
    ary(6) = "Density"
    ary(7) = "Material"
    Asm.SetCurrentFormat ary

Dim opath: opath = KCL.GetPath(KCL.getVbaDir & "\" & "oTemp")

opath = KCL.GetPath(KCL.getVbaDir)

para(0) = opath

Set txt = CATIA.SystemService.ExecuteScript(opath, 1, "getbom.CATscript", CATMain, para())

End Sub
