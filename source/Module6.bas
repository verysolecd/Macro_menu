Attribute VB_Name = "Module6"
Private Declare PtrSafe Function OpenClipboard Lib "user32" (ByVal hwnd As LongPtr) As Long
Private Declare PtrSafe Function EmptyClipboard Lib "user32" () As Long
Private Declare PtrSafe Function CloseClipboard Lib "user32" () As Long
Private Declare PtrSafe Function SetClipboardData Lib "user32" (ByVal wFormat As Long, ByVal hMem As LongPtr) As LongPtr


Option Explicit

' 声明 Windows API 函数，用于将窗口置于前台，这比 Visible=False/True 的技巧更可靠
#If VBA7 Then
    Private Declare PtrSafe Function SetForegroundWindow Lib "user32" (ByVal hwnd As LongPtr) As Long
#Else
    Private Declare Function SetForegroundWindow Lib "user32" (ByVal hwnd As Long) As Long
#End If

'''
' 激活或打开一个指定路径的文件资源管理器窗口。
' 如果该路径的窗口已经打开，则将其激活并显示在最前端；
' 如果没有打开，则新建一个窗口并打开该路径。
'
' @param targetPath 目标文件夹的完整路径，例如 "C:\Users"
'''
Public Sub ActivateOrOpenExplorer(ByVal targetPath As String)
    Dim shellApp As Object
    Dim shellWindow As Object
    Dim isWindowFound As Boolean
    
    isWindowFound = False
    
    ' 规范化路径，移除末尾的斜杠，确保比较的一致性
    If Right(targetPath, 1) = "\" And Len(targetPath) > 3 Then
        targetPath = Left(targetPath, Len(targetPath) - 1)
    End If

    On Error Resume Next ' 临时忽略错误，因为访问某些窗口属性可能会失败
    
    Set shellApp = CreateObject("Shell.Application")
    
    If shellApp Is Nothing Then Exit Sub
    
    For Each shellWindow In shellApp.Windows
        ' 关键检查：
        ' 1. 确保 shellWindow 对象有效
        ' 2. 检查窗口是否由 "explorer.exe" 创建，从而排除浏览器等其他窗口
        If Not shellWindow Is Nothing And LCase(Right(shellWindow.FullName, 11)) = "explorer.exe" Then
            
            ' 获取窗口当前的路径并规范化
            Dim currentPath As String
            currentPath = shellWindow.Document.Folder.Self.Path
            If Right(currentPath, 1) = "\" And Len(currentPath) > 3 Then
                currentPath = Left(currentPath, Len(currentPath) - 1)
            End If
            
            ' 不区分大小写地比较路径
            If LCase(currentPath) = LCase(targetPath) Then
                ' 找到了匹配的窗口
                isWindowFound = True
                
                ' 激活窗口
                shellWindow.Visible = True
                SetForegroundWindow shellWindow.hwnd
                
                Exit For ' 找到后即可退出循环
            End If
        End If
    Next shellWindow
    
    On Error GoTo 0 ' 恢复正常的错误处理
    
    ' 如果循环结束后仍未找到匹配的窗口，则打开一个新窗口
    If Not isWindowFound Then
        shellApp.Open targetPath
    End If
    
    Set shellWindow = Nothing
    Set shellApp = Nothing
End Sub

' === 如何使用 ===
' 创建一个模块并调用 ActivateOrOpenExplorer
Sub TestMyCode()
    ' 请将下面的路径替换为您想测试的实际文件夹路径
    Dim myPath As String
    myPath = "C:\Windows"
    
    Debug.Print "正在尝试激活或打开: " & myPath
    ActivateOrOpenExplorer myPath
End Sub