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
'    ' ��ȡCATIAӦ�ó���ʵ��
'    Set CATIA = GetObject(, "CATIA.Application")
'    Set activeWindow = CATIA.activeWindow
'
'    ' ��ʱ�ļ�·��
'    tempFilePath = Environ$("TEMP") & "\temp_screenshot.png"
'
'    ' �������ڽ�ͼ
'    activeWindow.ActiveViewer.CaptureToFile tempFilePath, 0, True ' 0��ʾPNG��ʽ
'
'    ' ���Ƶ�������
'    CopyImageToClipboard tempFilePath
'
'    ' ����
'    Kill tempFilePath
'
'    Set activeWindow = Nothing
'    Set CATIA = Nothing
'End Sub
'
'Private Sub CopyImageToClipboard(ByVal imagePath As String)
'    ' ʹ��Windows Script Host����ShellӦ�ó������
'    Dim objShell As Object
'    Set objShell = CreateObject("WScript.Shell")
'
'    ' ʹ��PowerShell�����ͼƬ��������
'    Dim psCommand As String
'    psCommand = "powershell -command ""Add-Type -AssemblyName System.Windows.Forms; [System.Windows.Forms.Clipboard]::SetImage([System.Drawing.Image]::FromFile('"" & imagePath & ""'))"""
'
'    objShell.Run psCommand, 0, True
'
'    Set objShell = Nothing
'End Sub
