Attribute VB_Name = "Module3"
Option Explicit

' ==============================================================================
' 函数名称: CheckClashBetweenTwoProducts
' 功能: 校核两个零部件/组件之间的干涉或间隙
' 输入:
'   - p1 (Product): 第一个产品
'   - p2 (Product): 第二个产品
'   - clearanceVal (Double): 安全间隙值(mm)。
'       * 设为 0 表示仅检查硬干涉和接触。
'       * 设为 >0 (例如 2.0) 表示检查 2mm 内的间隙，小于此距离视为干涉。
' 输出:
'   - String: 返回 "Interference" (硬干涉), "Contact" (接触), "Clearance Violation" (间隙违规) 或 "Safe" (安全)
'   - CATIA界面: 在结构树 Applications -> Clash 下生成分析对象
' ==============================================================================
Function CheckClashBetweenTwoProducts(p1, p2, Optional clearanceVal As Double = 0#) As String
    
    Dim doc As ProductDocument
    Set doc = CATIA.ActiveDocument
    
    Dim rootProd As Product
    Set rootProd = doc.Product

    Dim cClashes As Clashes
    On Error Resume Next
    Set cClashes = rootProd.GetTechnologicalObject("Clashes")
    On Error GoTo 0
    
    If cClashes Is Nothing Then
        MsgBox "无法获取Clashes对象，请检查是否拥有 SPA/DMU 许可证。", vbCritical
        CheckClashBetweenTwoProducts = "Error"
        Exit Function
    End If
    
    ' 2. 创建一个新的干涉分析
    Dim oClash As Clash
    Set oClash = cClashes.Add()
    
    ' 3. 设置计算类型：两组之间 (Between two selections)
    oClash.ComputationType = catClashComputationTypeBetweenTwo
    


    ' 4. 定义两组产品
     oClash.FirstGroup.AddExplicit p2  '
    
    oClash.FirstGroup.AddExplicit p1
   
    ' 5. 设置干涉类型和间隙值
    If clearanceVal > 0 Then
        ' 检查间隙模式
        oClash.InterferenceType = catClashInterferenceTypeClearance
        oClash.Clearance = clearanceVal
    Else
        ' 仅接触/干涉模式
        oClash.InterferenceType = catClashInterferenceTypeContact
    End If
    
    ' 6. 运行计算
    oClash.Compute
    
    ' 7. 重命名树上的节点，方便用户识别
    oClash.Name = "Check_" & p1.PartNumber & "_VS_" & p2.PartNumber
    
    ' 8. 分析结果逻辑
    ' 遍历所有冲突，判断最严重的干涉级别
    ' 优先级: Clash (硬干涉) > Contact (接触) > Clearance (间隙不足) > Safe
    
    Dim resultStr As String
    resultStr = "Safe"
    
    If oClash.Conflicts.count > 0 Then
        Dim i As Integer
        Dim oConflict As Conflict
        
        ' 预设状态
        Dim hasClash As Boolean: hasClash = False
        Dim hasContact As Boolean: hasContact = False
        Dim hasClearanceIssue As Boolean: hasClearanceIssue = False
        
        For i = 1 To oClash.Conflicts.count
            Set oConflict = oClash.Conflicts.item(i)
            
            If oConflict.Type = catConflictTypeClash Then
                hasClash = True
                Exit For ' 发现硬干涉，这是最严重的，直接退出循环
            ElseIf oConflict.Type = catConflictTypeContact Then
                hasContact = True
            ElseIf oConflict.Type = catConflictTypeClearance Then
                hasClearanceIssue = True
            End If
        Next i
        
        ' 根据优先级判定最终结果
        If hasClash Then
            resultStr = "Interference"   ' 存在硬干涉
        ElseIf hasContact Then
            resultStr = "Contact"        ' 存在接触 (如果 clearanceVal=0，这通常不算硬干涉，视需求而定)
        ElseIf hasClearanceIssue Then
            resultStr = "Clearance Violation" ' 间隙小于设定值
        End If
    End If
    
    CheckClashBetweenTwoProducts = resultStr

End Function

' ==============================================================================
' 测试过程：运行此 Sub 来测试上述函数
' ==============================================================================
Sub Test_Clash_Check()
    ' 1. 环境检查
    Dim doc As Document
    Set doc = CATIA.ActiveDocument
    
    Dim root As Product
    Set root = doc.Product
       
    ' 3. 获取前两个组件进行测试 (实际使用中你可以修改为 Selection 获取)
    Dim prod1 As Product
    Dim prod2 As Product
    Set prod1 = root.Products.item(1)
    Set prod2 = root.Products.item(2)
    
    ' 4. 调用函数
    ' 示例：检查 prod1 和 prod2，要求最小间隙为 2.0mm
    Dim checkResult As String
    checkResult = CheckClashBetweenTwoProducts(prod1, prod2, 3#)
    
    ' 5. 输出结果
    Dim msg As String
    msg = "校核完成！" & vbCrLf & vbCrLf
    msg = msg & "组件 1: " & prod1.PartNumber & vbCrLf
    msg = msg & "组件 2: " & prod2.PartNumber & vbCrLf
    msg = msg & "结果状态: " & checkResult & vbCrLf & vbCrLf
    msg = msg & "请在结构树底部的 'Applications -> Clash' 中查看详细的可视化结果。"
    
    MsgBox msg, vbInformation, "干涉检查结果"
    
End Sub
