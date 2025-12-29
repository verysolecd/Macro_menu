' ============================================
' 模块功能: 终极安全保存系统 (纯净版 - 无属性注入)
'
' [工作原理 - 纯文件属性流]
' 1. 初始化: 所有文件在硬盘上被设为 [只读] -> 无法保存
' 2. 解锁: 选中文件在硬盘上被设为 [可写] -> 唯一可保存的特征
' 3. 保存: 扫描所有打开文档，只保存那些 [可写] 的文件
'
' [包含宏命令]
' 1. OpenProductReadOnly  : [推荐] 按钮打开文件 -> 自动全锁
' 2. InitializeSafetyLock : [补救] 拖拽打开后 ->手动点此全锁
' 3. UnlockSelection      : [编辑] 选中零件 -> 解锁 (仅修改文件属性，不改模型)
' 4. CheckAndSaveUnlocked : [保存] 强制保存 (仅保存可写文件)
' ============================================
Option Explicit

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
           "提示: 现在的解锁机制非常纯净，不会在属性中写入任何标记。", vbInformation, "安全模式启动"
End Sub

' ============================================
' 2. 补救上锁 (✅ 拖拽打开文件后，请立即点此按钮!)
' 功能: 弥补拖拽打开无法自动上锁的缺陷，强制加上硬盘锁
' ============================================
Sub InitializeSafetyLock()
    LockAllFiles_Internal
    MsgBox "【安全系统已激活】" & vbCrLf & _
           "所有文件都已强制设为硬盘只读。" & vbCrLf & _
           "系统已就绪，等待您的解锁指令。", vbInformation, "手动锁定完成"
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
' 3. 解锁功能 (选中产品 -> 变为可写)
' 功能: 仅移除硬盘文件的"只读"属性，不做任何其他修改
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
            
            ' 获取文档对象
            Set doc = prod.ReferenceProduct.Parent
            If Not doc Is Nothing Then
                docPath = doc.FullName
                If docPath <> "" Then
                    ' 【关键】移除只读属性 (设为 vbNormal)
                    ' 这就是唯一的"解锁标记"，没有Param，非常干净
                    SetAttr docPath, vbNormal
                    
                    If Err.Number = 0 Then
                        unlockedCount = unlockedCount + 1
                        Debug.Print "已解锁(变更为可写): " & doc.Name
                    End If
                End If
            End If
        End If
        Err.Clear
    Next i
    
    If unlockedCount > 0 Then
        MsgBox "成功解锁 " & unlockedCount & " 个文件。" & vbCrLf & _
               "现在它是硬盘上唯一的'可写'文件，可以被保存宏识别。", vbInformation
    Else
        MsgBox "未能解锁文件。请确保选中的是有效的产品节点且已保存过。", vbExclamation
    End If
End Sub

' ============================================
' 4. 强制保存 (核心更新 - 纯净识别)
' 功能: 扫描所有文件，只保存那些硬盘属性为"可写"的文档
' ============================================
Sub CheckAndSaveUnlocked()
    Dim catApp As CATIA.Application
    Set catApp = CATIA.Application
    
    Dim docs As Documents
    Set docs = catApp.Documents
    
    If docs.Count = 0 Then Exit Sub
    
    Dim doc As Document
    Dim i As Integer
    Dim savedCount As Integer
    savedCount = 0
    Dim attr As VbFileAttribute
    
    On Error Resume Next
    
    ' 遍历所有打开的文档
    For i = 1 To docs.Count
        Set doc = docs.Item(i)
        
        ' 仅处理已保存过的、有路径的文件
        If doc.FullName <> "" Then
            ' 获取文件属性
            attr = GetAttr(doc.FullName)
            
            ' ??? 核心判断 ??? 
            ' 检查文件是否包含 vbReadOnly 属性
            ' 如果 (attr And vbReadOnly) 为 0，说明没有只读属性 -> 它是可写的(Unlocked)
            If (attr And vbReadOnly) = 0 Then
                
                ' === 核心逻辑: 突破只读保存 ===
                If doc.ReadOnly Then
                    ' 如果是CATIA只读模式打开的，必须用SaveAs覆盖原文件
                    ' 因为我们这一步确认了硬盘是可写的，所以SaveAs会成功
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
                Debug.Print "跳过(只读保护中): " & doc.Name
            End If
        End If
    Next i
    
    If savedCount > 0 Then
        MsgBox "保存完成!" & vbCrLf & "成功保存 " & savedCount & " 个文件(这些文件在硬盘上是可写状态)。", vbInformation
    Else
        MsgBox "没有保存任何文件。" & vbCrLf & "原因：所有文件在硬盘上都是只读的(未解锁)。", vbExclamation
    End If
End Sub