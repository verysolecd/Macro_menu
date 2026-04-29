Attribute VB_Name = "Module6"
Private Declare PtrSafe Function OpenClipboard Lib "user32" (ByVal hwnd As LongPtr) As Long
Private Declare PtrSafe Function EmptyClipboard Lib "user32" () As Long
Private Declare PtrSafe Function CloseClipboard Lib "user32" () As Long
Private Declare PtrSafe Function SetClipboardData Lib "user32" (ByVal wFormat As Long, ByVal hMem As LongPtr) As LongPtr
'Sub main()
'
'
'Call CaptureAndCopyToClipboard
'End Sub
'Public Sub CaptureAndCopyToClipboard()
'    Dim CATIA As Object
'    Dim activeWindow As Object
'    Dim tempFilePath As String
'
'    ' 获取CATIA应用程序实例
'    Set CATIA = GetObject(, "CATIA.Application")
'    Set activeWindow = CATIA.activeWindow
'
'    ' 临时文件路径
'    tempFilePath = Environ$("TEMP") & "\temp_screenshot.png"
'
'    ' 捕获活动窗口截图
'    activeWindow.ActiveViewer.CaptureToFile tempFilePath, 0, True ' 0表示PNG格式
'
'    ' 复制到剪贴板
'    CopyImageToClipboard tempFilePath
'
'    ' 清理
'    Kill tempFilePath
'
'    Set activeWindow = Nothing
'    Set CATIA = Nothing
'End Sub
'
'Private Sub CopyImageToClipboard(ByVal imagePath As String)
'    ' 使用Windows Script Host创建Shell应用程序对象
'    Dim objShell As Object
'    Set objShell = CreateObject("WScript.Shell")
'
'    ' 使用PowerShell命令复制图片到剪贴板
'    Dim psCommand As String
'    psCommand = "powershell -command ""Add-Type -AssemblyName System.Windows.Forms; [System.Windows.Forms.Clipboard]::SetImage([System.Drawing.Image]::FromFile('"" & imagePath & ""'))"""
'
'    objShell.Run psCommand, 0, True
'
'    Set objShell = Nothing
'End Sub
