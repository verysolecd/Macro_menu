Function AddToMarkdown() As Boolean
    Dim filePath As String
    Dim userInput As String
    Dim fso As Object
    Dim ts As Object
    
    ' 设置Markdown文件路径，可根据需要修改
    filePath = "C:\Users\YourName\Documents\笔记.md"
    
    ' 获取用户输入
    userInput = InputBox("请输入要添加到Markdown文件的内容：", "添加内容")
    
    ' 检查用户是否取消了输入
    If userInput = "" Then
        AddToMarkdown = False
        Exit Function
    End If
    
    On Error GoTo ErrorHandler
    
    ' 创建文件系统对象
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    ' 以追加模式打开文件，若文件不存在则创建
    Set ts = fso.OpenTextFile(filePath, 8, True)
    
    ' 写入用户输入内容，并添加换行符
    ts.WriteLine userInput
    
    ' 关闭文件
    ts.Close
    
    AddToMarkdown = True
    Exit Function
    
ErrorHandler:
    MsgBox "写入文件时出错: " & Err.Description, vbExclamation
    AddToMarkdown = False
    If Not ts Is Nothing Then ts.Close
End Function



Sub ExportSTPAndCompress()
    On Error Resume Next
    
    ' 获取当前活动文档
    Dim oDoc As Document
    Set oDoc = CATIA.ActiveDocument
    
    ' 检查是否为产品文档
    If TypeName(oDoc) <> "ProductDocument" Then
        MsgBox "请确保当前打开的是产品文档!", vbExclamation
        Exit Sub
    End If
    
    ' 获取产品名称（去除扩展名）
    Dim productName As String
    productName = Replace(oDoc.Name, ".CATProduct", "")
    
    ' 截取第一个下划线前的前缀（如DX11_DDD → DX11）
    Dim prefix As String
    Dim underscorePos As Integer
    underscorePos = InStr(productName, "_")
    
    If underscorePos > 0 Then
        prefix = Left(productName, underscorePos - 1)
    Else
        prefix = productName
    End If
    
    ' 获取当前日期（格式化为YYMMDD）
    Dim currentDate As String
    currentDate = Right(Year(Date), 2) & _
                  Right("0" & Month(Date), 2) & _
                  Right("0" & Day(Date), 2)
    
    ' 选择输出文件夹
    Dim ShellApp, folbrowser, folderoutput
    Set ShellApp = CreateObject("Shell.Application")
    Set folbrowser = ShellApp.BrowseForFolder(0, "选择STP输出文件夹", 16, 17)
    
    If folbrowser Is Nothing Then
        MsgBox "未选择输出文件夹，操作取消!", vbExclamation
        Exit Sub
    End If
    
    folderoutput = folbrowser.Self.path
    
    ' 构建STP文件名（例如：DX11_231005.stp）
    Dim stpPath As String
    stpPath = folderoutput & "\" & prefix & "_" & currentDate & ".stp"
    
    ' 导出为STP
    oDoc.ExportData stpPath, "stp"
    
    If Err.Number <> 0 Then
        MsgBox "STP导出失败：" & Err.Description, vbCritical
        Exit Sub
    End If
    
    ' 构建7-Zip压缩包路径（例如：DX11_231005.7z）
    Dim zipPath As String
    zipPath = folderoutput & "\" & prefix & "_" & currentDate & ".7z"
    
    ' 调用7-Zip命令行工具进行压缩
    Dim shell, cmd
    Set shell = CreateObject("WScript.Shell")
    
    ' 7-Zip命令：a = 添加到压缩包，-t7z = 指定7z格式，-mx=9 = 最高压缩率
    cmd = """C:\Program Files\7-Zip\7z.exe"" a -t7z -mx=9 """ & zipPath & """ """ & stpPath & """"
    
    ' 执行命令
    Dim result
    result = shell.Run(cmd, 0, True)  ' 0 = 隐藏窗口，True = 等待命令执行完成
    
    If result <> 0 Then
        MsgBox "7-Zip压缩失败！请确保7-Zip已正确安装。", vbCritical
    Else
        MsgBox "STP导出并压缩成功：" & vbCrLf & stpPath & vbCrLf & zipPath, vbInformation
        
        ' 可选：删除原始STP文件（保留压缩包）
        ' If FileExists(stpPath) Then Kill stpPath
    End If
    
    ' 释放对象
    Set shell = Nothing
    Set folbrowser = Nothing
    Set ShellApp = Nothing
    Set oDoc = Nothing
    
    On Error GoTo 0
End Sub

' 辅助函数：检查文件是否存在
Function FileExists(filePath)
    Dim fso
    Set fso = CreateObject("Scripting.FileSystemObject")
    FileExists = fso.FileExists(filePath)
    Set fso = Nothing
End Function






Sub ExportSTPAndZipWithWin11()
    On Error Resume Next
    
    ' 获取当前活动文档
    Dim oDoc As Document
    Set oDoc = CATIA.ActiveDocument
    
    ' 检查是否为产品文档
    If TypeName(oDoc) <> "ProductDocument" Then
        MsgBox "请确保当前打开的是产品文档!", vbExclamation
        Exit Sub
    End If
    
    ' 获取产品名称（去除扩展名）
    Dim productName As String
    productName = Replace(oDoc.Name, ".CATProduct", "")
    
    ' 截取第一个下划线前的前缀（如DX11_DDD → DX11）
    Dim prefix As String
    Dim underscorePos As Integer
    underscorePos = InStr(productName, "_")
    
    If underscorePos > 0 Then
        prefix = Left(productName, underscorePos - 1)
    Else
        prefix = productName
    End If
    
    ' 获取当前日期（格式化为YYMMDD）
    Dim currentDate As String
    currentDate = Right(Year(Date), 2) & _
                  Right("0" & Month(Date), 2) & _
                  Right("0" & Day(Date), 2)
    
    ' 选择输出文件夹
    Dim ShellApp, folbrowser, folderoutput
    Set ShellApp = CreateObject("Shell.Application")
    Set folbrowser = ShellApp.BrowseForFolder(0, "选择STP输出文件夹", 16, 17)
    
    If folbrowser Is Nothing Then
        MsgBox "未选择输出文件夹，操作取消!", vbExclamation
        Exit Sub
    End If
    
    folderoutput = folbrowser.Self.path
    
    ' 构建STP文件名（例如：DX11_231005.stp）
    Dim stpPath As String
    stpPath = folderoutput & "\" & prefix & "_" & currentDate & ".stp"
    
    ' 导出为STP
    oDoc.ExportData stpPath, "stp"
    
    If Err.Number <> 0 Then
        MsgBox "STP导出失败：" & Err.Description, vbCritical
        Exit Sub
    End If
    
    ' 构建ZIP压缩包路径（例如：DX11_231005.zip）
    Dim zipPath As String
    zipPath = folderoutput & "\" & prefix & "_" & currentDate & ".zip"
    
    ' 创建空ZIP文件（Windows 11原生支持）
    CreateEmptyZipFile zipPath
    
    ' 等待ZIP文件创建完成
    Dim fso
    Set fso = CreateObject("Scripting.FileSystemObject")
    Dim waitTime As Integer
    waitTime = 0
    
    Do While Not fso.FileExists(zipPath) And waitTime < 10
        WScript.Sleep 500  ' 等待0.5秒
        waitTime = waitTime + 1
    Loop
    
    If Not fso.FileExists(zipPath) Then
        MsgBox "创建ZIP文件失败！", vbCritical
        Exit Sub
    End If
    
    ' 将STP文件添加到ZIP压缩包
    Dim sourceFile, destinationZip
    Set sourceFile = ShellApp.NameSpace(stpPath).Items
    Set destinationZip = ShellApp.NameSpace(zipPath)
    
    If Not destinationZip Is Nothing Then
        destinationZip.CopyHere sourceFile, 4  ' 4 = 不显示确认对话框
        
        ' 等待压缩完成（非阻塞）
        WScript.Sleep 2000  ' 等待2秒（可根据文件大小调整）
        
        MsgBox "STP导出并压缩成功：" & vbCrLf & stpPath & vbCrLf & zipPath, vbInformation
        
        ' 可选：删除原始STP文件（保留压缩包）
        ' If fso.FileExists(stpPath) Then fso.DeleteFile stpPath
    Else
        MsgBox "无法访问ZIP文件！", vbCritical
    End If
    
    ' 释放对象
    Set fso = Nothing
    Set destinationZip = Nothing
    Set sourceFile = Nothing
    Set folbrowser = Nothing
    Set ShellApp = Nothing
    Set oDoc = Nothing
    
    On Error GoTo 0
End Sub

' 创建空ZIP文件的辅助函数
Sub CreateEmptyZipFile(zipFilePath)
    Dim fso, tempFile
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    ' 创建临时文件写入ZIP文件头
    Set tempFile = fso.CreateTextFile(zipFilePath, True)
    tempFile.Write "PK" & Chr(5) & Chr(6) & String(18, Chr(0))
    tempFile.Close
    
    Set tempFile = Nothing
    Set fso = Nothing
End Sub

Sub ExportSTPAndZipWithPowerShell()
    On Error Resume Next
    
    ' 获取当前活动文档
    Dim oDoc As Document
    Set oDoc = CATIA.ActiveDocument
    
    ' 检查是否为产品文档
    If TypeName(oDoc) <> "ProductDocument" Then
        MsgBox "请确保当前打开的是产品文档!", vbExclamation
        Exit Sub
    End If
    
    ' 获取产品名称（去除扩展名）
    Dim productName As String
    productName = Replace(oDoc.Name, ".CATProduct", "")
    
    ' 截取第一个下划线前的前缀（如DX11_DDD → DX11）
    Dim prefix As String
    Dim underscorePos As Integer
    underscorePos = InStr(productName, "_")
    
    If underscorePos > 0 Then
        prefix = Left(productName, underscorePos - 1)
    Else
        prefix = productName
    End If
    
    ' 获取当前日期（格式化为YYMMDD）
    Dim currentDate As String
    currentDate = Right(Year(Date), 2) & _
                  Right("0" & Month(Date), 2) & _
                  Right("0" & Day(Date), 2)
    
    ' 选择输出文件夹
    Dim ShellApp, folbrowser, folderoutput
    Set ShellApp = CreateObject("Shell.Application")
    Set folbrowser = ShellApp.BrowseForFolder(0, "选择STP输出文件夹", 16, 17)
    
    If folbrowser Is Nothing Then
        MsgBox "未选择输出文件夹，操作取消!", vbExclamation
        Exit Sub
    End If
    
    folderoutput = folbrowser.Self.path
    
    ' 构建STP文件名（例如：DX11_231005.stp）
    Dim stpPath As String
    stpPath = folderoutput & "\" & prefix & "_" & currentDate & ".stp"
    
    ' 导出为STP
    oDoc.ExportData stpPath, "stp"
    
    If Err.Number <> 0 Then
        MsgBox "STP导出失败：" & Err.Description, vbCritical
        Exit Sub
    End If
    
    ' 构建ZIP压缩包路径（例如：DX11_231005.zip）
    Dim zipPath As String
    zipPath = folderoutput & "\" & prefix & "_" & currentDate & ".zip"
    
    ' 使用PowerShell命令压缩文件
    Dim shell, cmd, result
    Set shell = CreateObject("WScript.Shell")
    
    ' PowerShell命令：使用最快压缩级别(-CompressionLevel Fastest)
    cmd = "powershell -Command ""Compress-Archive -Path '""" & stpPath & """' -DestinationPath '""" & zipPath & """' -CompressionLevel Fastest -Force"""
    
    ' 执行命令并获取返回值
    result = shell.Run(cmd, 0, True)
    
    If result <> 0 Then
        MsgBox "PowerShell压缩失败！请确保PowerShell版本不低于5.0。", vbCritical
    Else
        ' 验证ZIP文件是否存在
        Dim fso
        Set fso = CreateObject("Scripting.FileSystemObject")
        
        If fso.FileExists(zipPath) Then
            MsgBox "STP导出并压缩成功：" & vbCrLf & stpPath & vbCrLf & zipPath, vbInformation
        Else
            MsgBox "压缩完成但未找到ZIP文件！", vbCritical
        End If
        
        Set fso = Nothing
    End If
    
    ' 释放对象
    Set shell = Nothing
    Set folbrowser = Nothing
    Set ShellApp = Nothing
    Set oDoc = Nothing
    
    On Error GoTo 0
End Sub



