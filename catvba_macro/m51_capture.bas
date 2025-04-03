Attribute VB_Name = "mCaptureClipboard"
'{GP:4}
'{Ep:CaptureToClipboard}
'{Caption:��ͼ��������}
'{ControlTipText:����ǰCATIA��ͼ��ͼ���Ƶ�������}

' ��Ҫ����Windows API����
Private Declare Function OpenClipboard Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function EmptyClipboard Lib "user32" () As Long
Private Declare Function CloseClipboard Lib "user32" () As Long
Private Declare Function SetClipboardData Lib "user32" (ByVal wFormat As Long, ByVal hMem As Long) As Long
Private Declare Function CopyImage Lib "user32" (ByVal handle As Long, ByVal un1 As Long, ByVal n1 As Long, ByVal n2 As Long, ByVal un2 As Long) As Long

Const CF_BITMAP = 2

Sub CaptureToClipboard()
    On Error GoTo ErrorHandler
    
    ' ��ȡCATIAӦ�úͻ����
    Dim catia As Application
    Set catia = CATIA
    
    If catia.ActiveWindow Is Nothing Then
        MsgBox "û�л���ڿɽ�ͼ", vbExclamation
        Exit Sub
    End If
    
    ' ��ȡ��ǰ��ͼ
    Dim viewer As Viewer
    Set viewer = catia.ActiveWindow.ActiveViewer
    
    ' ��ʱ�ļ�·��
    Dim tempFile As String
    tempFile = Environ("TEMP") & "\CATIA_temp_capture.bmp"
    
    ' �ȱ���Ϊ��ʱBMP�ļ�
    viewer.CaptureToFile 0, tempFile ' 0 = catCaptureFormatBMP
    
    ' ��BMP�ļ����Ƶ�������
    Dim hBitmap As Long
    hBitmap = LoadImage(0, tempFile, 0, 0, 0, &H10) ' LR_LOADFROMFILE
    
    If hBitmap <> 0 Then
        OpenClipboard 0
        EmptyClipboard
        SetClipboardData CF_BITMAP, hBitmap
        CloseClipboard
        MsgBox "��ͼ�Ѹ��Ƶ�������", vbInformation
    Else
        MsgBox "�޷���ͼ���Ƶ�������", vbExclamation
    End If
    
    ' ɾ����ʱ�ļ�
    Kill tempFile
    
    Exit Sub
    
ErrorHandler:
    MsgBox "��ͼʧ�ܣ�" & Err.Description, vbCritical
End Sub
