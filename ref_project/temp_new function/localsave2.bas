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



'优化函数

'----------------------------------------------------------------------------
' Macro: CATIA_BatchSaveAllProductsByLevel.catvbs
' 功能：批量保存CATIA装配体所有子产品/零件（按层级从下到上、自动去重、区分文件后缀）
' 优化点：自定义保存路径、完善错误处理、用户交互提示、语法规范化
'----------------------------------------------------------------------------
Option Explicit '强制声明变量，避免语法错误

' ========== 自定义类：保存产品层级和对象信息 ==========
Class clsSaveInfo
    Public level As Integer      ' 产品在装配体中的层级（根产品为1级）
    Public prod As Product       ' 产品对象（Product/Part）
    Public partNumber As String  ' 零件号（冗余存储，方便后续扩展）
End Class

' ========== 主函数：程序入口 ==========
Sub CATMain()
    ' 1. 初始化配置
    Dim originalAlertState As Boolean
    originalAlertState = CATIA.DisplayFileAlerts ' 保存原始弹窗状态
    CATIA.DisplayFileAlerts = False              ' 关闭CATIA文件操作弹窗
    
    ' 2. 检查当前文档类型
    If CATIA.ActiveDocument.Type <> "Product" Then
        MsgBox "错误：当前打开的不是装配体（Product）文档！", vbCritical + vbOKOnly, "CATIA批量保存"
        CATIA.DisplayFileAlerts = originalAlertState
        Exit Sub
    End If
    
    ' 3. 让用户选择保存文件夹
    Dim savePath As String
    savePath = SelectSaveFolder()
    If savePath = "" Then ' 用户取消选择
        MsgBox "操作已取消：未选择保存文件夹", vbInformation + vbOKOnly, "CATIA批量保存"
        CATIA.DisplayFileAlerts = originalAlertState
        Exit Sub
    End If
    
    ' 4. 核心逻辑：收集装配体所有产品信息 + 按层级保存
    On Error GoTo ErrorHandler ' 全局错误捕获
    Dim rootProd As Product
    Set rootProd = CATIA.ActiveDocument.Product
    
    ' 创建字典存储待保存产品（Key=零件号，Value=clsSaveInfo对象）
    Dim docsToSave As Scripting.Dictionary
    Set docsToSave = New Scripting.Dictionary
    
    Dim maxLevel As Integer
    maxLevel = 0
    
    ' 递归遍历装配体，收集所有产品信息（去重）
    Call SlurpAssembly(1, rootProd, docsToSave, maxLevel)
    
    ' 按层级从深到浅保存（先保存底层零件，再保存上层装配体）
    Call SaveProductsByLevel(docsToSave, maxLevel, savePath)
    
    ' 5. 操作完成提示
    MsgBox "批量保存完成！" & vbCrLf & "保存路径：" & savePath, vbInformation + vbOKOnly, "CATIA批量保存"
    
Cleanup:
    ' 恢复原始配置
    CATIA.DisplayFileAlerts = originalAlertState
    Set docsToSave = Nothing
    Set rootProd = Nothing
    Exit Sub
    
ErrorHandler:
    MsgBox "保存失败：" & Err.Description & vbCrLf & "错误代码：" & Err.Number, vbCritical + vbOKOnly, "CATIA批量保存"
    Resume Cleanup
End Sub

' ========== 辅助函数：弹出文件夹选择窗口 ==========
Function SelectSaveFolder() As String
    Const WINDOW_HANDLE = 0
    Const NO_OPTIONS = 0
    Dim objShell As Object, objFolder As Object, objFolderItem As Object
    
    Set objShell = CreateObject("Shell.Application")
    ' 弹出文件夹选择窗口（默认路径为桌面，更贴近用户习惯）
    Set objFolder = objShell.BrowseForFolder(WINDOW_HANDLE, "请选择保存文件夹：", NO_OPTIONS, CreateObject("WScript.Shell").SpecialFolders("Desktop"))
    
    If Not objFolder Is Nothing Then
        Set objFolderItem = objFolder.Self
        SelectSaveFolder = objFolderItem.Path
    Else
        SelectSaveFolder = "" ' 用户取消选择
    End If
    
    Set objShell = Nothing
    Set objFolder = Nothing
    Set objFolderItem = Nothing
End Function

' ========== 辅助函数：递归遍历装配体，收集产品信息（去重） ==========
Sub SlurpAssembly(ByVal currentLevel As Integer, ByRef aProd As Product, ByRef allDocs As Scripting.Dictionary, ByRef maxLevel As Integer)
    ' 更新最大层级（记录最深的子产品）
    If currentLevel > maxLevel Then
        maxLevel = currentLevel
    End If
    
    ' 去重逻辑：零件号不存在时才添加（避免同一零件多次保存）
    Dim prodPartNumber As String
    prodPartNumber = Trim(aProd.PartNumber) ' 去除首尾空格，避免命名异常
    
    If prodPartNumber = "" Then
        prodPartNumber = "Unnamed_" & Replace(CreateObject("Scriptlet.TypeLib").GUID, "-", "") ' 给无零件号的产品生成唯一名称
    End If
    
    If Not allDocs.Exists(prodPartNumber) Then
        Dim info As clsSaveInfo
        Set info = New clsSaveInfo
        info.level = currentLevel
        Set info.prod = aProd
        info.partNumber = prodPartNumber
        allDocs.Add prodPartNumber, info ' 添加到字典
    End If
    
    ' 递归遍历子产品
    Dim i As Integer
    For i = 1 To aProd.Products.Count
        Dim subProd As Product
        Set subProd = aProd.Products.Item(i)
        Call SlurpAssembly(currentLevel + 1, subProd, allDocs, maxLevel)
    Next
    
    Set subProd = Nothing
End Sub

' ========== 辅助函数：按层级从深到浅保存所有产品 ==========
Sub SaveProductsByLevel(ByRef docsToSave As Scripting.Dictionary, ByVal maxLevel As Integer, ByVal savePath As String)
    Dim currentLevel As Integer
    currentLevel = maxLevel
    
    Do While currentLevel >= 1
        Dim toRemove As Collection
        Set toRemove = New Collection
        
        ' 遍历当前层级的所有产品并保存
        Dim key As Variant
        For Each key In docsToSave.Keys
            Dim info As clsSaveInfo
            Set info = docsToSave(key)
            
            If info.level = currentLevel Then
                ' 区分文件类型，拼接正确后缀
                Dim fileSuffix As String
                If TypeName(info.prod.ReferenceProduct.Parent) = "PartDocument" Then
                    fileSuffix = ".CATPart"
                Else
                    fileSuffix = ".CATProduct"
                End If
                
                ' 拼接完整保存路径
                Dim fullPath As String
                fullPath = savePath & "\" & info.partNumber & fileSuffix
                
                ' 保存文档（跳过已打开且路径相同的文件，避免重复保存）
                Dim targetDoc As Document
                Set targetDoc = info.prod.ReferenceProduct.Parent
                If targetDoc.FullName <> fullPath Then
                    targetDoc.SaveAs fullPath
                End If
                
                toRemove.Add key ' 标记已保存的产品，后续从字典删除
            End If
        Next
        
        ' 移除已保存的产品
        Dim i As Integer
        For i = 1 To toRemove.Count
            docsToSave.Remove toRemove(i)
        Next
        
        currentLevel = currentLevel - 1 ' 层级递减，处理上一层
    Loop
    
    Set toRemove = Nothing
End Sub