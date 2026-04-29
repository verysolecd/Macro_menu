' 
'----------------------------------------------------------------------------
' Macro: CatiaV5-AllProductsLocalSaving.catvbs
''----------------------------------------------------------------------------
Sub CATMain()
    CATIA.DisplayFileAlerts = False
    'BrowseForFile
    Const WINDOW_HANDLE = 0
    Const NO_OPTIONS = 0
    Set objShell = CreateObject("Shell.Application")
    Set objFolder = objShell.BrowseForFolder _
    (WINDOW_HANDLE, "Select a folder:", NO_OPTIONS, "C:\")
    Set objFolderItem = objFolder.Self
    objPath = objFolderItem.Path
'Get the root of the CATProduct
    Dim rootPrd As Product
    Set rootPrd = CATIA.ActiveDocument.Product
'此处增加产品类型检查canexecute
'Recursive function localSaveAs
    localSaveAs rootPrd, objPath
CATIA.DisplayFileAlerts = True
End Sub
Function localSaveAs(rootPrdItem, objPath)
    Dim subRootProduct As Product
    For Each subRootProduct In rootPrdItem.Products
        toSave = subRootProduct.ReferenceProduct.Parent.Name
        CATIA.Documents.Item(toSave).SaveAs (objPath & "\" & i & toSave)
        localSaveAs subRootProduct, objPath
    Next
End Function
'----------------------------------------------------------------------------
' Macro: CatiaV5-AllProductsLocalSaving.catvbs
' 功能：按零件号将总成和所有子产品另存为到指定文件夹
'----------------------------------------------------------------------------
Sub CATMain()
    On Error Resume Next
    CATIA.DisplayFileAlerts = False
    ' 选择保存目录
    Const WINDOW_HANDLE = 0
    Const NO_OPTIONS = 0
    Set objShell = CreateObject("Shell.Application")
    Set objFolder = objShell.BrowseForFolder _
    (WINDOW_HANDLE, "选择保存文件夹:", NO_OPTIONS, "C:\")
    If objFolder Is Nothing Then
        MsgBox "用户取消了操作"
        Exit Sub
    End If
    Set objFolderItem = objFolder.Self
    objPath = objFolderItem.Path
    ' 获取根产品
    Dim rootPrd As Product
    Set rootPrd = CATIA.ActiveDocument.Product
    ' 检查文档类型
    If CATIA.ActiveDocument.Type <> "Product" Then
        MsgBox "当前文档不是产品文档，无法执行此操作"
        Exit Sub
    End If
    ' 递归保存所有产品
    localSaveAs rootPrd, objPath
    MsgBox "所有产品已成功保存到: " & objPath
    CATIA.DisplayFileAlerts = True
End Sub
Function localSaveAs(rootPrdItem, objPath)
    On Error Resume Next
    Dim subRootProduct As Product
    Dim i As Integer
    i = 1  ' 初始化计数器
    ' 保存当前产品
    Dim currentDoc As Document
    Set currentDoc = rootPrdItem.ReferenceProduct.Parent
    If Not currentDoc Is Nothing Then
        Dim partNumber As String
        partNumber = rootPrdItem.PartNumber
        ' 使用零件号作为文件名
        'toSave = subRootProduct.ReferenceProduct.Parent.Name
        Dim fileName As String
        fileName = partNumber & ".CATProduct"
        ' 完整文件路径
        Dim fullPath As String
        fullPath = objPath & "\" & fileName
        ' 保存当前产品
        currentDoc.SaveAs fullPath
    End If
    ' 递归保存子产品
    For Each subRootProduct In rootPrdItem.Products
        localSaveAs subRootProduct, objPath
    Next
End Function
----------Class clsSaveInfo definition--------------
Public level As Integer
Public prod As Product
-----------------(module definition)--------------- 
Option Explicit
Sub CATMain()
    CATIA.DisplayFileAlerts = False
    'get the root product
    Dim rootProd As Product
    Set rootProd = CATIA.ActiveDocument.Product
    'make a dictionary to track product structure
    Dim docsToSave As Scripting.Dictionary
    Set docsToSave = New Scripting.Dictionary
    'some parameters
    Dim level As Integer
    Dim maxLevel As Integer
    'read the assembly
    level = 0
    Call slurp(level, rootProd, docsToSave, maxLevel)
    Dim i
    Dim kx As String
    Dim info As clsSaveInfo
    Do Until docsToSave.count = 0
        Dim toRemove As Collection
        Set toRemove = New Collection
        For i = 0 To docsToSave.count - 1
           kx = docsToSave.keys(i)
           Set info = docsToSave.item(kx)
           If info.level = maxLevel Then
                Dim suffix As String
               If TypeName(info.prod) = "Part" Then
                    suffix = ".CATPart"
               Else
                    suffix = ".CATProduct"
                End If
                Dim partProd As Product
                Set partProd = info.prod
                Dim partDoc As Document
                Set partDoc = partProd.ReferenceProduct.Parent
                partDoc.SaveAs ("C:\Temp\" & partProd.partNumber & suffix)
                toRemove.add (kx)
            End If
        Next
     'remove the saved products from the dictionary
        For i = 1 To toRemove.count
            docsToSave.Remove (toRemove.item(i))
        Next
        'decrement the level we are looking for
        maxLevel = maxLevel - 1
    Loop
End Sub
Sub slurp(ByVal level As Integer, ByRef aProd As Product, ByRef allDocs As Scripting.Dictionary, ByRef maxLevel As Integer)
'increment the level
    level = level + 1
'track the max level
    If level > maxLevel Then maxLevel = level
 'see if the part is already in the save list, if not add it
    If allDocs.Exists(aProd.partNumber) = False Then
        Dim info As clsSaveInfo
        Set info = New clsSaveInfo
        info.level = level
        Set info.prod = aProd
        Call allDocs.add(aProd.partNumber, info)
    End If
'slurp up children
    Dim i
    For i = 1 To aProd.products.count
        Dim subProd As Product
        Set subProd = aProd.products.item(i)
        Call slurp(level, subProd, allDocs, maxLevel)
    Next
End Sub