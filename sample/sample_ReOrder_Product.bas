'Attribute VB_Name = "sample_ReOrder_Product"

'{GP:11}
'{Caption:产品排序}
'{ControlTipText:产品排序}
'{BackColor:16744703}

Option Explicit

Sub CATMain()

    If Not CanExecute("ProductDocument") Then Exit Sub
    

    Dim ProDoc As ProductDocument: Set ProDoc = CATIA.ActiveDocument
    Dim Pros As Products: Set Pros = ProDoc.Product.Products
    If Pros.Count < 2 Then Exit Sub
    

    Dim AssyMode As AsmConstraintSettingAtt
    Set AssyMode = CATIA.SettingControllers.Item("CATAsmConstraintSettingCtrl")
    Dim OriginalMode As CatAsmPasteComponentMode
    OriginalMode = AssyMode.PasteComponentMode
    

    AssyMode.PasteComponentMode = catPasteWithCstOnCopyAndCut
    

    Dim Names: Set Names = Get_SortedNames(Pros)
    

    Dim Sel As Selection: Set Sel = ProDoc.Selection
    Dim Itm As Variant
    
    CATIA.HSOSynchronized = False
    
    Sel.Clear
    For Each Itm In Names
        Sel.Add Pros.Item(Itm)
    Next
    Sel.Cut
    

    With Sel
        .Clear
        .Add Pros
        .Paste
        .Clear
    End With
    
    CATIA.HSOSynchronized = True
    

    AssyMode.PasteComponentMode = OriginalMode
    ProDoc.Product.Update
End Sub

Private Function Get_SortedNames(ByVal Pros As Products) As Object
    Dim Lst As Object
    Set Lst = KCL.InitLst()
    
    Dim Pro As Product
    For Each Pro In Pros
        Lst.Add Pro.Name
    Next
    
    Lst.Sort
    
    Set Get_SortedNames = Lst
End Function
