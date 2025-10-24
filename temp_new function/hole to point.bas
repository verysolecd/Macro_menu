Sub CATMain()
    Set oDoc = CATIA.ActiveDocument
    Set oPart = oDoc.part
    Set HSF = oPart.HybridShapeFactory
    '======= 要求选择body
    Dim imsg, filter(0)
    imsg = "请选择body"
    filter(0) = "Body"
    Dim obdy
    Set obdy = KCL.SelectElement(imsg, filter).Value
    Set targethb = oPart.HybridBodies.Add()
    targethb.Name = "extracted points"
    If Not obdy Is Nothing Then
            Set holeBody = obdy
            For Each Hole In holeBody.Shapes
            If TypeOf Hole Is Hole Then
                Set skt = Hole.Sketch
                Set Pt = HSF.AddNewPointCoord(0, 0, 0)
                Set ref = oPart.CreateReferenceFromObject(skt)
                Pt.PtRef = ref
                Pt.Name = "Pt_" & i
                targethb.AppendHybridShape Pt
                oPart.InWorkObject = Pt
                oPart.Update
                i = i + 1
            End If
        Next
            MsgBox "完成：" & i & "个点", vbInformation
    End If

End Sub


' API函数声明
Public Declare Function FindWindow Lib "user32" alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long 
Public Declare Function SetForegroundWindow Lib "user32" (ByVal hwnd As Long) As Integer 

Sub WaitForCATIADisassembleAndContinue()
    ' 变量声明
    Dim confirmationWindowHandle As Long 
    Dim progressWindowHandle As Long 
    Dim shellObject As Object 
    ' 创建Shell对象用于发送键盘命令
    Set shellObject = CreateObject("Wscript.Shell") 
    
    ' 第一阶段：等待分解确认窗口出现
    Dim waitCounter As Integer
    waitCounter = 0
    
    ' 等待确认窗口（通常是"确认"或"OK"窗口）
    Do While waitCounter < 300  ' 最多等待30秒
        ' 尝试查找常见的CATIA确认窗口标题
        confirmationWindowHandle = FindWindow(vbNullString, "确认")  ' 中文版CATIA
        If confirmationWindowHandle = 0 Then
            confirmationWindowHandle = FindWindow(vbNullString, "OK")  ' 英文版CATIA
        End If
        If confirmationWindowHandle = 0 Then
            confirmationWindowHandle = FindWindow(vbNullString, "Disassemble")  ' 分解对话框
        End If
        
        DoEvents() 
        
        If confirmationWindowHandle <> 0 Then 
            ' 确认窗口已找到，将其置前并发送确认键
            SetForegroundWindow(confirmationWindowHandle) 
            
            ' 等待一小段时间确保窗口完全激活
            Application.Wait Now + TimeValue("00:00:01")
            
            ' 发送回车键确认
            shellObject.SendKeys "{ENTER}"
            
            ' 也可以尝试空格键或Tab+回车组合
            ' shellObject.SendKeys "{SPACE}"
            ' shellObject.SendKeys "{TAB}{ENTER}"
            
            Exit Do
        End If
        
        waitCounter = waitCounter + 1
    Loop
    
    If confirmationWindowHandle = 0 Then
        MsgBox "错误：未找到分解确认窗口，请手动确认或检查窗口标题"
        Exit Sub
    End If
    
    ' 第二阶段：等待进度条窗口出现并完成
    waitCounter = 0
    
    ' 等待进度条窗口出现
    Do While waitCounter < 600  ' 最多等待60秒
        ' 尝试查找进度条窗口（CATIA通常有"进度"或"Progress"窗口）
        progressWindowHandle = FindWindow(vbNullString, "进度")  ' 中文版
        If progressWindowHandle = 0 Then
            progressWindowHandle = FindWindow(vbNullString, "Progress")  ' 英文版
        End If
        If progressWindowHandle = 0 Then
            progressWindowHandle = FindWindow(vbNullString, "Processing")  ' 处理中窗口
        End If
        
        DoEvents()
        
        If progressWindowHandle <> 0 Then
            ' 进度条窗口已找到，等待它完成
            Exit Do
        End If
        
        waitCounter = waitCounter + 1
    Loop
    
    If progressWindowHandle = 0 Then
        ' 可能没有进度条窗口，直接等待几秒后继续
        MsgBox "未检测到进度条窗口，等待5秒后继续..."
        Application.Wait Now + TimeValue("00:00:05")
    Else
        ' 第三阶段：等待进度条窗口消失（表示分解完成）
        waitCounter = 0
        
        Do While waitCounter < 1200  ' 最多等待2分钟
            DoEvents()
            
            ' 检查进度条窗口是否还存在
            progressWindowHandle = FindWindow(vbNullString, "进度")
            If progressWindowHandle = 0 Then
                progressWindowHandle = FindWindow(vbNullString, "Progress")
            End If
            If progressWindowHandle = 0 Then
                progressWindowHandle = FindWindow(vbNullString, "Processing")
            End If
            
            If progressWindowHandle = 0 Then
                ' 进度条窗口已消失，分解完成
                Exit Do
            End If
            
            waitCounter = waitCounter + 1
        Loop
        
        If progressWindowHandle <> 0 Then
            MsgBox "警告：进度条窗口可能未正常关闭，但将继续执行后续代码"
        End If
    End If
    
    ' 第四阶段：分解完成，执行您的后续代码
    Call ExecuteYourSubsequentCode()
    
    MsgBox "CATIA分解命令已完成，后续代码已执行"
End Sub

' 您的后续代码放在这里
Sub ExecuteYourSubsequentCode()
    ' 这里放置您希望在分解完成后执行的代码
    
    ' 示例1：继续其他CATIA操作
    Dim catia As Object
    On Error Resume Next
    Set catia = GetObject(, "CATIA.Application")
    If Not catia Is Nothing Then
        ' 您的CATIA后续操作代码
        ' 例如：选择其他元素、创建新特征等
    End If
    
    ' 示例2：记录操作日志
    Debug.Print "分解操作完成于: " & Now()
    
    ' 示例3：更新界面状态
    ' YourForm.ProgressBar.Value = 100
    ' YourForm.StatusLabel.Caption = "分解完成"
    
    ' 示例4：调用其他宏或函数
    ' Call YourNextMacro()
    
    ' 请根据您的实际需求修改这里的代码
End Sub

' 增强版本：自动检测窗口标题
Sub WaitForCATIADisassembleEnhanced()
    Dim shellObject As Object
    Set shellObject = CreateObject("Wscript.Shell")
    
    ' 检测当前CATIA的窗口标题模式
    Dim windowTitles As Variant
    windowTitles = DetectCATIAWindowTitles()
    
    ' 等待确认窗口
    If WaitForWindowAndSendKey(windowTitles(0), "{ENTER}", 30) Then
        ' 等待进度条窗口完成
        If WaitForWindowCompletion(windowTitles(1), 120) Then
            ' 执行后续代码
            Call ExecuteYourSubsequentCode()
            MsgBox "分解操作成功完成"
        Else
            MsgBox "进度条窗口可能异常，但将继续执行后续代码"
            Call ExecuteYourSubsequentCode()
        End If
    Else
        MsgBox "未能找到确认窗口，请手动操作"
    End If
End Sub

' 检测CATIA窗口标题
Function DetectCATIAWindowTitles() As Variant
    Dim titles(1) As String
    
    ' 常见的CATIA窗口标题
    titles(0) = "确认"  ' 确认窗口标题
    titles(1) = "进度"  ' 进度窗口标题
    
    ' 您可以在这里添加更多的窗口标题检测逻辑
    DetectCATIAWindowTitles = titles
End Function

' 通用窗口等待和按键发送函数
Function WaitForWindowAndSendKey(windowTitle As String, keyToSend As String, maxWaitSeconds As Integer) As Boolean
    Dim windowHandle As Long
    Dim shellObject As Object
    Dim waitCounter As Integer
    
    Set shellObject = CreateObject("Wscript.Shell")
    waitCounter = 0
    
    Do While waitCounter < maxWaitSeconds * 10  ' 转换为十分之一秒
        windowHandle = FindWindow(vbNullString, windowTitle)
        DoEvents()
        
        If windowHandle <> 0 Then
            SetForegroundWindow(windowHandle)
            Application.Wait Now + TimeValue("00:00:01")
            shellObject.SendKeys keyToSend
            WaitForWindowAndSendKey = True
            Exit Function
        End If
        
        waitCounter = waitCounter + 1
    Loop
    
    WaitForWindowAndSendKey = False
End Function

' 等待窗口完成函数
Function WaitForWindowCompletion(windowTitle As String, maxWaitSeconds As Integer) As Boolean
    Dim windowHandle As Long
    Dim waitCounter As Integer
    
    waitCounter = 0
    
    ' 先等待窗口出现
    Do While waitCounter < maxWaitSeconds * 10
        windowHandle = FindWindow(vbNullString, windowTitle)
        DoEvents()
        
        If windowHandle <> 0 Then Exit Do
        waitCounter = waitCounter + 1
    Loop
    
    If windowHandle = 0 Then Return True  ' 可能没有进度窗口
    
    ' 等待窗口消失
    waitCounter = 0
    Do While waitCounter < maxWaitSeconds * 10
        DoEvents()
        windowHandle = FindWindow(vbNullString, windowTitle)
        
        If windowHandle = 0 Then
            WaitForWindowCompletion = True
            Exit Function
        End If
        
        waitCounter = waitCounter + 1
    Loop
    
    WaitForWindowCompletion = False
End Function