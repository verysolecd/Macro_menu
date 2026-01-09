Attribute VB_Name = "Module5"
'函数库


''==遍历递归=============================

Sub recurAyo(ayo)
    Dim colls: Set itm = ayo.Products
    For Each itm In colls
        Call recurFunc(itm)
    Next

    If ayo.Products.count > 0 Then
            For Each ctm In ayo.Products
                Call recurAyo(ctm)
             Next
    End If
End Sub

''==获取父级=============================
