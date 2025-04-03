Attribute VB_Name = "mCaptureClipboard"
'{GP:4}
'{Ep:CaptureToClipboard}
'{Caption:截图到剪贴板}
'{ControlTipText:将当前CATIA视图截图复制到剪贴板}

' 需要声明Windows API函数
Private Declare Function OpenClipboard Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function EmptyClipboard Lib "user32" () As Long
Private Declare Function CloseClipboard Lib "user32" () As Long
Private Declare Function SetClipboardData Lib "user32" (ByVal wFormat As Long, ByVal hMem As Long) As Long
Private Declare Function CopyImage Lib "user32" (ByVal handle As Long, ByVal un1 As Long, ByVal n1 As Long, ByVal n2 As Long, ByVal un2 As Long) As Long

Const CF_BITMAP = 2

Sub CaptureToClipboard()
    On Error GoTo ErrorHandler
    
    ' 获取CATIA应用和活动窗口
    Dim catia As Application
    Set catia = CATIA
    
    If catia.ActiveWindow Is Nothing Then
        MsgBox "没有活动窗口可截图", vbExclamation
        Exit Sub
    End If
    
    ' 获取当前视图
    Dim viewer As Viewer
    Set viewer = catia.ActiveWindow.ActiveViewer
    
    ' 临时文件路径
    Dim tempFile As String
    tempFile = Environ("TEMP") & "\CATIA_temp_capture.bmp"
    
    ' 先保存为临时BMP文件
    viewer.CaptureToFile 0, tempFile ' 0 = catCaptureFormatBMP
    
    ' 将BMP文件复制到剪贴板
    Dim hBitmap As Long
    hBitmap = LoadImage(0, tempFile, 0, 0, 0, &H10) ' LR_LOADFROMFILE
    
    If hBitmap <> 0 Then
        OpenClipboard 0
        EmptyClipboard
        SetClipboardData CF_BITMAP, hBitmap
        CloseClipboard
        MsgBox "截图已复制到剪贴板", vbInformation
    Else
        MsgBox "无法将图像复制到剪贴板", vbExclamation
    End If
    
    ' 删除临时文件
    Kill tempFile
    
    Exit Sub
    
ErrorHandler:
    MsgBox "截图失败：" & Err.Description, vbCritical
End Sub
