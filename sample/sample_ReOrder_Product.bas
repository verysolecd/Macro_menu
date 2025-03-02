' Attribute VB_Name = "sample_ReOrder_Product"
' VBA 示例：重新排序产品 版本 0.0.1，使用 'KCL0.0.12'，作者：Kantoku
' 可以对产品树中的项目进行重新排序
'{GP:11}
'{标题: 重新排序}
'{控件提示文本: 可以对产品树中的项目进行重新排序}
'{背景颜色: 16744703}
'{字体大小: 10.5}
Option Explicit
Sub CATMain()
    ' 检查是否可以执行
    If Not CanExecute("ProductDocument") Then Exit Sub
    
    ' 获取文档
    Dim ProDoc As ProductDocument: Set ProDoc = CATIA.ActiveDocument
    Dim Pros As Products: Set Pros = ProDoc.Product.Products
    If Pros.Count < 2 Then Exit Sub
    
    ' 获取装配模式
    Dim AssyMode As AsmConstraintSettingAtt
    Set AssyMode = CATIA.SettingControllers.Item("CATAsmConstraintSettingCtrl")
    Dim OriginalMode As CatAsmPasteComponentMode
    OriginalMode = AssyMode.PasteComponentMode
    
    ' 更改装配模式
    AssyMode.PasteComponentMode = catPasteWithCstOnCopyAndCut
    
    ' 获取排序后的名称
    Dim Names: Set Names = Get_SortedNames(Pros)
    
    ' 选择操作
    Dim Sel As Selection: Set Sel = ProDoc.Selection
    Dim Itm As Variant
    
    CATIA.HSOSynchronized = False
    
    Sel.Clear
    For Each Itm In Names
        Sel.Add Pros.Item(Itm)
    Next
    Sel.Cut
    
    ' 粘贴操作
    With Sel
        .Clear
        .Add Pros
        .Paste
        .Clear
    End With
    
    CATIA.HSOSynchronized = True
    
    ' 恢复装配模式并更新
    AssyMode.PasteComponentMode = OriginalMode
    ProDoc.Product.Update
End Sub
' 获取排序后的产品名称列表
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