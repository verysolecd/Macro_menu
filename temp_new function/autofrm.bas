' ============================================
' 模块功能: 终极安全保存系统 (全流程接管版)
'
' [包含宏命令]
' 1. OpenProductReadOnly  : [推荐] 选择文件并以绝对安全模式打开（自动上双重锁）
' 2. InitializeSafetyLock : [补救] 拖拽打开文件后，点此按钮进行手动上锁
' 3. UnlockSelection      : [编辑] 选中零件 -> 解锁权限
' 4. CheckAndSaveUnlocked : [保存] 强制保存修改成果
'
' [场景说明]
' - 方式A: 使用 OpenProductReadOnly 打开 -> 自动全锁 (最安全)
' - 方式B: 拖拽文件打开 -> 必须立即点击 InitializeSafetyLock -> 补上硬盘锁 (同样安全)
' ============================================
Option Explicit

' 解锁标记参数名
Const UNLOCK_FLAG_NAME = "Is_Unlocked"

' ============================================
' 1. 安全打开入口 (标准做法)
' 功能: 弹出对话框选择文件 -> 只读打开 -> 自动上锁
' ============================================
Sub OpenProductReadOnly()
    Dim catApp As CATIA.Application
    Set catApp = CATIA.Application
    
    ' 1. 弹出文件选择框
    Dim filePath As String
    filePath = catApp.FileSelectionBox("请选择要安全打开的产品", "*.CATProduct", CatFileSelectionModeOpen)
    
    If filePath = "" Then Exit Sub ' 用户取消
    
    On Error Resume Next
    
    ' 2. 以只读模式打开文档 (ReadOnly:=True)
    Dim doc As Document
    Set doc = catApp.Documents.Open(filePath, True)
    
    If Err.Number <> 0 Then
        MsgBox "打开文件失败: " & Err.Description, vbCritical
        Err.Clear
        Exit Sub
    End If
    
    ' 3. 自动上硬盘锁
    LockAllFiles_Internal
    
    MsgBox "文件已安全打开!" & vbCrLf & _
           "状态: [Session只读] + [硬盘只读]" & vbCrLf & _
           "需要编辑时，请使用解锁按钮。", vbInformation, "安全模式启动"
End Sub

' ============================================
' 辅助过程: 内部调用的上锁逻辑
' ============================================
Sub LockAllFiles_Internal()
    On Error Resume Next
    Dim catApp As CATIA.Application
    Set catApp = CATIA.Application
    Dim docs As Documents
    Set docs = catApp.Documents
    Dim i As Integer
    Dim doc As Document
    
    For i = 1 To docs.Count
        Set doc = docs.Item(i)
        If doc.FullName <> "" Then
            ' 强制设为只读
            SetAttr doc.FullName, vbReadOnly
            Err.Clear
        End If
    Next i
End Sub

' ============================================
' 2. 补救上锁 (✅ 拖拽打开文件后，请立即点此按钮!)
' 功能: 弥补拖拽打开无法自动上锁的缺陷，强制加上硬盘锁
' ============================================
Sub InitializeSafetyLock()
    LockAllFiles_Internal
    MsgBox "【安全系统已激活】" & vbCrLf & _
           "所有文件都已强制设为硬盘只读。" & vbCrLf & _
           "提示：即使您是拖拽打开的，现在也已处于保护状态。" & vbCrLf & _
           "任何原生保存操作都会失败，直到您手动解锁。", vbInformation, "手动锁定完成"
End Sub

' ============================================
' 3. 解锁功能 (选中产品 -> 授予保存权限)
' 功能: 
'   1. 移除硬盘文件的"只读"属性 (确保SaveAs能写入)
'   2. 添加 "Is_Unlocked" 标记 (作为允许保存的凭证)
' ============================================
Sub UnlockSelection()
    Dim catApp As CATIA.Application
    Set catApp = CATIA.Application
    
    Dim sel As Selection
    Set sel = catApp.ActiveDocument.Selection
    
    If sel.Count = 0 Then
        MsgBox "请先选择要解锁的产品或零件!", vbExclamation
        Exit Sub
    End If
    
    Dim i As Integer
    Dim prod As Product
    Dim doc As Document
    Dim docPath As String
    Dim unlockedCount As Integer
    unlockedCount = 0
    
    On Error Resume Next
    
    For i = 1 To sel.Count
        If TypeName(sel.Item(i).Value) = "Product" Then
            Set prod = sel.Item(i).Value
            
            ' 1. 尝试添加解锁标记(参数)
            MarkProductAsUnlocked prod
            
            ' 2. 尝试获取文档并去除只读属性
            Set doc = prod.ReferenceProduct.Parent
            If Not doc Is Nothing Then
                docPath = doc.FullName
                If docPath <> "" Then
                    ' 【关键】移除只读属性 (设为 vbNormal)
                    ' 这使得该文件成为全场唯一可写的文件
                    SetAttr docPath, vbNormal
                    
                    If Err.Number = 0 Then
                        unlockedCount = unlockedCount + 1
                        Debug.Print "已解锁(权限): " & doc.Name
                    End If
                End If
            End If
        End If
        Err.Clear
    Next i
    
    If unlockedCount > 0 Then
        MsgBox "成功解锁 " & unlockedCount & " 个文件。" & vbCrLf & _
               "现在它已变身'可写'状态，之后可用保存宏进行保存。", vbInformation
    Else
        MsgBox "未能解锁文件。请确保选中的是有效的产品节点。", vbExclamation
    End If
End Sub

' ============================================
' 4. 强制保存 (核心更新)
' 功能: 扫描解锁标记，智能突破只读锁定进行保存
' ============================================
Sub CheckAndSaveUnlocked()
    Dim catApp As CATIA.Application
    Set catApp = CATIA.Application
    
    Dim docs As Documents
    Set docs = catApp.Documents
    
    If docs.Count = 0 Then Exit Sub
    
    Dim doc As Document
    Dim prodDoc As ProductDocument
    Dim rootProd As Product
    Dim i As Integer
    Dim savedCount As Integer
    savedCount = 0
    
    On Error Resume Next
    
    ' 遍历所有打开的文档
    For i = 1 To docs.Count
        Set doc = docs.Item(i)
        
        ' 仅处理Product和Part文档
        If TypeName(doc) = "ProductDocument" Or TypeName(doc) = "PartDocument" Then
            
            ' 获取根节点
            Set rootProd = Nothing
            If TypeName(doc) = "ProductDocument" Then
                Set rootProd = doc.Product
            ElseIf TypeName(doc) = "PartDocument" Then
                Set rootProd = doc.Product
            End If
            
            If Not rootProd Is Nothing Then
                ' 检查是否有"Is_Unlocked"标记
                If IsProductUnlocked(rootProd) Then
                    
                    ' 确保硬盘文件可写(防止Unlock后又被人手动设为只读)
                    If doc.FullName <> "" Then SetAttr doc.FullName, vbNormal
                    
                    ' === 核心逻辑: 突破只读保存 ===
                    If doc.ReadOnly Then
                        ' 如果是CATIA只读模式打开的，必须用SaveAs覆盖原文件
                        doc.SaveAs doc.FullName
                    Else
                        ' 如果是正常模式，直接Save
                        doc.Save
                    End If
                    
                    If Err.Number = 0 Then
                        savedCount = savedCount + 1
                        Debug.Print "已保存: " & doc.Name
                    Else
                        Debug.Print "保存失败: " & doc.Name & " - " & Err.Description
                        Err.Clear
                    End If
                Else
                    Debug.Print "跳过(未解锁): " & doc.Name
                End If
            End If
        End If
    Next i
    
    If savedCount > 0 Then
        MsgBox "保存完成!" & vbCrLf & "成功保存 " & savedCount & " 个已解锁文件。", vbInformation
    Else
        MsgBox "没有保存任何文件。" & vbCrLf & "原因可能是：没有找到带有 [" & UNLOCK_FLAG_NAME & "] 标记的文件。", vbExclamation
    End If
End Sub

' ---------------------------------------------------------
' 辅助: 添加解锁标记
' ---------------------------------------------------------
Function MarkProductAsUnlocked(targetProd As Product)
    On Error Resume Next
    Dim params As Parameters
    Set params = targetProd.UserRefProperties
    Dim p As Parameter
    ' 检查是否已存在
    Err.Clear
    Set p = params.Item(UNLOCK_FLAG_NAME)
    ' 不存在则创建
    If Err.Number <> 0 Then
        Set p = params.CreateBoolean(UNLOCK_FLAG_NAME, True)
    End If
    ' 确保为True
    p.ValuateFromString "True"
End Function

' ---------------------------------------------------------
' 辅助: 检查解锁标记
' ---------------------------------------------------------
Function IsProductUnlocked(targetProd As Product) As Boolean
    On Error Resume Next
    IsProductUnlocked = False
    Dim params As Parameters
    Set params = targetProd.UserRefProperties
    Dim p As Parameter
    Set p = params.Item(UNLOCK_FLAG_NAME)
    If Err.Number = 0 Then
        If p.Value = True Then IsProductUnlocked = True
    End If
End Function