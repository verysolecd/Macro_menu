
'{Gp:51}
'{Ep:LeftHand}
'{Caption: 我是标题}
'{ControlTipText: 检查零件文档中是否存在左手坐标系}
'{背景颜色: 33023}





Option Explicit
Private mFSO As Object
Sub CATMain()
    ' 检查文档
    If Not KCL.CanExecute("ProductDocument") Then Exit Sub
    
    Set mFSO = KCL.GetFSO()
    Dim msg As String
    
    ' 获取顶层对象
    Dim root As Products
    Set root = CATIA.ActiveDocument.Product.Products
    
    Dim docs As Collection
    Set docs = New Collection
    Set docs = GetRenameDoc(GetAllDoc(root, docs))
    
    If docs.Count < 1 Then
        msg = "没有需要修正的文档。"
        MsgBox msg, vbExclamation
        Exit Sub
    End If
    
    ' 确认
    msg = GetRenameListMsg(docs)
    If MsgBox(msg, vbYesNo + vbQuestion) = vbNo Then
        Exit Sub
    End If
    
    ' 重命名
    Call ExecRename(docs)
    
    ' 结束
    Set mFSO = Nothing
    MsgBox ("完成")
    
End Sub
Private Sub ExecRename( _
    ByVal docs As Collection)
    
    Dim doc As AnyObject
    Dim tmp As Variant
    For Each doc In docs
        tmp = GetFilename_PartNum(doc)
        doc.Product.PartNumber = tmp(0)
    Next
End Sub
    
Private Function GetRenameListMsg( _
    ByVal docs As Collection) As String
    
    Dim ary() As Variant
    ReDim ary(docs.Count)
    
    ary(0) = "是否将以下的零件编号改为文件名？" + vbCrLf + _
        "是否确认？"
        
    Dim tmp As Variant
    Dim doc As AnyObject
    Dim i As Long
    For i = 1 To docs.Count
        Set doc = docs.Item(i)
        tmp = GetFilename_PartNum(doc)
        ary(i) = tmp(1) + " → " + tmp(0)
    Next
    
    GetRenameListMsg = Join(ary, vbCrLf)
        
End Function
Private Function GetFilename_PartNum( _
    ByVal doc As Document) As Variant
    
    GetFilename_PartNum = Array( _
        mFSO.GetBaseName(doc.FullName), _
        doc.Product.PartNumber)
        
End Function
Private Function GetRenameDoc( _
    ByVal docs As Collection) As Collection
    
    Dim lst As Collection
    Set lst = New Collection
    
    Dim doc As AnyObject
    Dim tmp As Variant
    For Each doc In docs
        tmp = GetFilename_PartNum(doc)
        If tmp(0) <> tmp(1) Then
            Call lst.Add(doc)
        End If
    Next
    
    Set GetRenameDoc = lst
    
End Function
Private Function GetAllDoc( _
    ByVal prods As Products, _
    ByVal lst As Collection) As Collection
    
    Dim prod As Product
    For Each prod In prods
        Call lst.Add(prod.ReferenceProduct.Parent)
        Set lst = GetAllDoc(prod.Products, lst)
    Next
    
    Set GetAllDoc = lst
    
End Function