Sub CATMain()
    On Error GoTo ErrorHandler
    GetNextNode CATIA.ActiveDocument.Product
    Exit Sub
ErrorHandler:
    MsgBox "执行过程中出现错误: " & Err.Description, vbCritical, "错误"
End Sub

Sub GetNextNode(oCurrentProduct As Product)
    Dim oCurrentTreeNode As Product
    Dim StrNomenclature As String, StrDesignation As String, StrWindows As String, PartNumber As String
    Dim i As Integer
    
    ' Loop through every tree node for the current product+
    For i = 1 To oCurrentProduct.Products.Count
        Set oCurrentTreeNode = oCurrentProduct.Products.Item(i)
        PartNumber = oCurrentTreeNode.PartNumber
        
        ' Determine if the current node is a part, product or component
        If Right(oCurrentTreeNode.Name, 4) = "Part" Then
            PictureGet PartNumber
            MsgBox PartNumber & " is a part"
        ElseIf IsProduct(oCurrentTreeNode) Then
            PictureGet PartNumber
            MsgBox PartNumber & " is a product"
        Else
            PictureGet PartNumber
            MsgBox PartNumber & " is a component"
        End If
        
        ' if sub-nodes exist below the current tree node, call the sub recursively
        If oCurrentTreeNode.Products.Count > 0 Then
            GetNextNode oCurrentTreeNode
        End If
    Next i
End Sub

Function IsProduct(objCurrentProduct As Product) As Boolean
    Dim oTestProduct As ProductDocument
    Set oTestProduct = Nothing
    On Error Resume Next
    Set oTestProduct = CATIA.Documents.Item(objCurrentProduct.PartNumber & ".CATProduct")
    On Error GoTo 0
    IsProduct = Not oTestProduct Is Nothing
End Function

Function PictureGet(PartName As String, oCurrentProduct As Product) As String
    Dim ObjViewer3D As Viewer3D
    Set objViewer3D = CATIA.ActiveWindow.ActiveViewer
    
    Dim objCamera3D As Camera3D
    Set objCamera3D = CATIA.ActiveDocument.Cameras.Item(1)
    
    If PartName = "" Then
        MsgBox "No name was entered. Operation aborted.", vbExclamation, "Cancel"
    Else
        'turn off the spec tree
        Dim objSpecWindow As SpecsAndGeomWindow
        Set objSpecWindow = CATIA.ActiveWindow
        objSpecWindow.Layout = catWindowGeomOnly
        
        '=== 新增: 聚焦到当前组件 ===
        CATIA.ActiveDocument.Selection.Clear
        CATIA.ActiveDocument.Selection.Add oCurrentProduct
        objViewer3D.Reframe ' 这将使视图聚焦到选中的组件
        '=========================
        
        'Toggle Compass
        CATIA.StartCommand("Compass")
        
        'change background color to white
        Dim DBLBackArray(2)
        objViewer3D.GetBackgroundColor(dblBackArray)
        Dim dblWhiteArray(2)
        dblWhiteArray(0) = 1
        dblWhiteArray(1) = 1
        dblWhiteArray(2) = 1
        objViewer3D.PutBackgroundColor(dblWhiteArray)
        
        'file location to save image
        Dim fileloc As String
        fileloc = "C:\Temp\"
        
        Dim exten As String
        exten = ".jpg"
        
        Dim strName As String
        strName = fileloc & PartName & exten
        
        'clear selection for picture
        CATIA.ActiveDocument.Selection.Clear()
        
        'increase to fullscreen to obtain maximum resolution
        objViewer3D.FullScreen = True
        
        'take picture
        objViewer3D.CaptureToFile 4, strName
        
        '*******************RESET**********************
        objViewer3D.FullScreen = False
        objViewer3D.PutBackgroundColor(dblBackArray)
        objSpecWindow.Layout = catWindowSpecsAndGeom
        CATIA.StartCommand("Compass")
    End If
End Function