Dim currentDateTime As String
currentDateTime = Year(Date) & _
                  Right("0" & Month(Date), 2) & _
                  Right("0" & Day(Date), 2) & "_" & _
                  Right("0" & Hour(Now), 2) & _
                  Right("0" & Minute(Now), 2) & _
                  Right("0" & Second(Now), 2)
' 生成如：20231001_143045




截取文件名
Dim fileName
fileName = "DX11_DDD_20231005.stp"

' 查找第一个下划线的位置
Dim underscorePos
underscorePos = InStr(fileName, "_")

' 如果找到下划线，则截取前面的字符
Dim prefix
If underscorePos > 0 Then
    prefix = Left(fileName, underscorePos - 1)  ' 截取从左边开始到下划线前的字符
Else
    prefix = fileName  ' 如果没有下划线，返回原文件名
End If

MsgBox "截取结果: " & prefix  ' 输出: DX11




Sub ExportSTPWithPrefixAndDate()
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
        prefix = productName  ' 如果没有下划线，使用完整名称
    End If
    
    ' 获取当前日期（年份取后两位，格式化为YYMMDD）
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
    
    ' 构建带前缀和日期的STP文件名（例如：DX11_231005.stp）
    Dim outputPath As String
    outputPath = folderoutput & "\" & prefix & "_" & currentDate & ".stp"
    
    ' 导出为STP
    oDoc.ExportData outputPath, "stp"
    
    If Err.Number <> 0 Then
        MsgBox "导出失败：" & Err.Description, vbCritical
    Else
        MsgBox "STP导出成功：" & outputPath, vbInformation
    End If
    
    ' 释放对象
    Set folbrowser = Nothing
    Set ShellApp = Nothing
    Set oDoc = Nothing
    
    On Error GoTo 0
End Sub