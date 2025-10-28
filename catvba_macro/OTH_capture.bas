Attribute VB_Name = "OTH_capture"
'Attribute VB_Name = "mCaptureClipboard"
'{GP:64}
'{Ep:CaptureToClipboard}
'{Caption:��ͼ��������}
'{ControlTipText:����ǰCATIA��ͼ��ͼ���Ƶ�������}
'{BackColor:16744703}


' ��Ҫ����Windows API����
'#If VBA7 Then
'    Private Declare PtrSafe Function OpenClipboard Lib "user32" (ByVal hwnd As LongPtr) As LongPtr
'    Private Declare PtrSafe Function EmptyClipboard Lib "user32" () As LongPtr
'    Private Declare PtrSafe Function CloseClipboard Lib "user32" () As LongPtr
'    Private Declare PtrSafe Function SetClipboardData Lib "user32" (ByVal wFormat As Long, ByVal hMem As LongPtr) As LongPtr
'    Private Declare PtrSafe Function CopyImage Lib "user32" (ByVal handle As LongPtr, ByVal un1 As Long, ByVal n1 As Long, ByVal n2 As Long, ByVal un2 As Long) As LongPtr
'    Private Declare PtrSafe Function LoadImage Lib "user32" Alias "LoadImageA" (ByVal hInst As LongPtr, ByVal lpsz As String, ByVal un1 As Long, ByVal n1 As Long, ByVal n2 As Long, ByVal un2 As Long) As LongPtr
'#Else
'    Private Declare Function OpenClipboard Lib "user32" (ByVal hwnd As Long) As Long
'    Private Declare Function EmptyClipboard Lib "user32" () As Long
'    Private Declare Function CloseClipboard Lib "user32" () As Long
'    Private Declare Function SetClipboardData Lib "user32" (ByVal wFormat As Long, ByVal hMem As Long) As Long
'    Private Declare Function CopyImage Lib "user32" (ByVal handle As Long, ByVal un1 As Long, ByVal n1 As Long, ByVal n2 As Long, ByVal un2 As Long) As Long
'    Private Declare Function LoadImage Lib "user32" Alias "LoadImageA" (ByVal hInst As Long, ByVal lpsz As String, ByVal un1 As Long, ByVal n1 As Long, ByVal n2 As Long, ByVal un2 As Long) As Long
'#End If

Const CF_BITMAP = 2

Sub CaptureToClipboard()

MsgBox "û����"
'    On Error GoTo ErrorHandler
'
'    ' ��ȡCATIAӦ�úͻ����
'    Dim catia As Application
'    Set catia = catia
'
'    If catia.ActiveWindow Is Nothing Then
'        MsgBox "û�л���ڿɽ�ͼ", vbExclamation
'        Exit Sub
'    End If
'
'    ' ��ȡ��ǰ��ͼ
'    Dim viewer As viewer
'    Set viewer = catia.ActiveWindow.ActiveViewer
'
'    ' ��ʱ�ļ�·��
'    Dim tempFile As String
'    tempFile = Environ("TEMP") & "\CATIA_temp_capture.bmp"
'
'    ' �ȱ���Ϊ��ʱBMP�ļ�
'    viewer.CaptureToFile 0, tempFile ' 0 = catCaptureFormatBMP
'
'    ' ��BMP�ļ����Ƶ�������
'    Dim hBitmap As Long
'    hBitmap = LoadImage(0, tempFile, 0, 0, 0, &H10) ' LR_LOADFROMFILE
'
'    If hBitmap <> 0 Then
'        OpenClipboard 0
'        EmptyClipboard
'        SetClipboardData CF_BITMAP, hBitmap
'        CloseClipboard
'        MsgBox "��ͼ�Ѹ��Ƶ�������", vbInformation
'    Else
'        MsgBox "�޷���ͼ���Ƶ�������", vbExclamation
'    End If
'
'    ' ɾ����ʱ�ļ�
'    Kill tempFile
'
'    Exit Sub
'
'ErrorHandler:
'    MsgBox "��ͼʧ�ܣ�" & Err.Description, vbCritical
End Sub


Function PictureGet(PartName As String, oprd) As String
    Dim ObjViewer3D As Viewer3D
    Set ObjViewer3D = CATIA.ActiveWindow.ActiveViewer
    
    Dim objCamera3D As Camera3D
    Set objCamera3D = CATIA.ActiveDocument.Cameras.item(1)
    
    If PartName = "" Then
        MsgBox "No name was entered. Operation aborted.", vbExclamation, "Cancel"
    Else
        'turn off the spec tree
        Dim objSpecWindow As SpecsAndGeomWindow
        Set objSpecWindow = CATIA.ActiveWindow
        objSpecWindow.Layout = catWindowGeomOnly
        
        '=== ����: �۽�����ǰ��� ===
        CATIA.ActiveDocument.Selection.Clear
        CATIA.ActiveDocument.Selection.Add oprd
        ObjViewer3D.Reframe ' �⽫ʹ��ͼ�۽���ѡ�е����
        '=========================
        
        'Toggle Compass
        CATIA.StartCommand ("Compass")
        
        'change background color to white
        Dim DBLBackArray(2)
        ObjViewer3D.GetBackgroundColor (DBLBackArray)
        Dim dblWhiteArray(2)
        dblWhiteArray(0) = 1
        dblWhiteArray(1) = 1
        dblWhiteArray(2) = 1
        ObjViewer3D.PutBackgroundColor (dblWhiteArray)
        
        'file location to save image
        Dim fileloc As String
        fileloc = "C:\Temp\"
        
        Dim exten As String
        exten = ".jpg"
        
        Dim strName As String
        strName = fileloc & PartName & exten
        
        'clear selection for picture
        CATIA.ActiveDocument.Selection.Clear()
        
        'increase to fullscreen to obtain maximum resolution
        ObjViewer3D.FullScreen = True
        
        'take picture
        ObjViewer3D.CaptureToFile 4, strName
        
        '*******************RESET**********************
        ObjViewer3D.FullScreen = False
        ObjViewer3D.PutBackgroundColor (DBLBackArray)
        objSpecWindow.Layout = catWindowSpecsAndGeom
        CATIA.StartCommand ("Compass")
    End If
End Function
