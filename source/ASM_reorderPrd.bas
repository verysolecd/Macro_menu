Attribute VB_Name = "ASM_reorderPrd"
'Attribute VB_Name = "sample_ReOrder_Product"
'{GP:3}
'{Ep:CATMain}
'{Caption:产品排序}
'{ControlTipText:产品排序}
'{BackColor: }
Option Explicit
Sub CATMain()
    If Not CanExecute("ProductDocument") Then Exit Sub
    Dim prodoc As ProductDocument: Set prodoc = CATIA.ActiveDocument
    Dim Pros As Products: Set Pros = prodoc.Product.Products
    If Pros.count < 2 Then Exit Sub
    Dim AssyMode As AsmConstraintSettingAtt
    Set AssyMode = CATIA.SettingControllers.item("CATAsmConstraintSettingCtrl")
    Dim OriginalMode As CatAsmPasteComponentMode
    OriginalMode = AssyMode.PasteComponentMode
    AssyMode.PasteComponentMode = catPasteWithCstOnCopyAndCut
    Dim Names: Set Names = Get_SortedNames(Pros)
    Dim sel As Selection: Set sel = prodoc.Selection
    Dim itm As Variant
    CATIA.HSOSynchronized = False
    sel.Clear
    For Each itm In Names
        sel.Add Pros.item(itm)
    Next
    sel.Cut
    With sel
        .Clear
        .Add Pros
        .Paste
        .Clear
    End With
    CATIA.HSOSynchronized = True
    AssyMode.PasteComponentMode = OriginalMode
    prodoc.Product.Update
End Sub
Private Function Get_SortedNames(ByVal Pros As Products) As Object
    Dim lst As Object
    Set lst = KCL.InitLst()
    Dim Pro As Product
    For Each Pro In Pros
        lst.Add Pro.Name
    Next
    lst.Sort
    Set Get_SortedNames = lst
End Function
