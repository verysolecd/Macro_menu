Attribute VB_Name = "OTH_capture"
'Attribute VB_Name = "mCaptureClipboard"
'{GP:6}
'{Ep:CaptureTopath}
'{Caption:截图到文件夹}
'{ControlTipText:遍历产品并截图到文件夹}
'{BackColor:16744703}
' 需要声明Windows API函数
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
Option Explicit

' ==============================================================================
' Module: OTH_capture
' Description: Captures screenshots of CATIA products recursively.
' ==============================================================================

' --- Configuration Types ---
Private Type CaptureSettings
    BackgroundColor As Variant
    FocusFactor As Double
    ViewDirection As Variant
    RenderingMode As Integer
    AspectHeightRatio As Double
    ResolutionWidth As Long
End Type

' --- Constants ---
Private Const DEFAULT_FOCUS_FACTOR = 0.9
Private Const MDL_NAME As String = "OTH_capture"
' --- State Variables ---
Private m_Settings As CaptureSettings
Private m_ProcessedPN As Object ' Scripting.Dictionary
Private m_FirstImagePath As String

' ==============================================================================
' Public Entry Point
' ==============================================================================
Private Const mdlname As String = "OTH_capture"
Sub CaptureTopath()
    If Not KCL.CanExecute("ProductDocument,PartDocument") Then Exit Sub
    CATIA.StartCommand ("* iso")
    Dim response As VbMsgBoxResult
    response = MsgBox("如要截图，请等待ISO视角调整完毕后点击确认", vbYesNo + vbExclamation, "确认截图")
    If response <> vbYes Then Exit Sub
    CapPrd CATIA.ActiveDocument.Product
    
    If m_FirstImagePath <> "" Then
        KCL.SmartOPenPath KCL.ofParentPath(m_FirstImagePath)
    End If
End Sub
' ==============================================================================
' Core Execution Logic
' ==============================================================================
Sub CapPrd(oPrd)
'    On Error GoTo ErrorHandler
    If Not KCL.CanExecute("ProductDocument,PartDocument") Then Exit Sub
    If oPrd Is Nothing Then Exit Sub
    ' Setup Environment
     InitializeSettings
    CATIA.StartCommand ("Compass") ' Toggle Compass (Hide)
    CATIA.RefreshDisplay = False
    oPrd.ApplyWorkMode 3 ' DESIGN_MODE
    HideNonBody CATIA.ActiveDocument
    SetCAPDisplay
    Set m_ProcessedPN = KCL.InitDic
    Dim tempPath As String    ' Prepare Output Folder
    tempPath = KCL.GetPath(KCL.getVbaDir & "\oTemp"): KCL.ClearDir tempPath
    g_Picpath = tempPath
    CaptureRecursive oPrd, tempPath     ' Start Recursive Capture
    Set m_ProcessedPN = Nothing
    RecoverDisplay
    CATIA.StartCommand ("Compass") ' Toggle Compass (Restore)
    CATIA.RefreshDisplay = True
    Exit Sub
ErrorHandler:
    CATIA.RefreshDisplay = True
    MsgBox "Error in CapPrd: " & Err.Description, vbCritical, MDL_NAME
End Sub

Private Sub CaptureRecursive(targetPrd, ByVal folderPath As String)
    Dim viewer As viewer: Set viewer = CATIA.ActiveWindow.ActiveViewer
    Dim partNumber As String: partNumber = targetPrd.partNumber
    If Not m_ProcessedPN.Exists(partNumber) Then
        Dim imgFile As String
        imgFile = folderPath & "\" & partNumber & ".jpg"
        viewer.CaptureToFile 5, imgFile ' 5 = catCaptureFormatJPEG
        If m_FirstImagePath = "" Then m_FirstImagePath = imgFile
        m_ProcessedPN.Add partNumber, 1
    End If
    Dim children As Products: Set children = targetPrd.Products
    If children.count = 0 Then Exit Sub
    Dim sel As Selection: Set sel = CATIA.ActiveDocument.Selection
    Dim visp: Set visp = sel.VisProperties
    ' Hide all children first (Performance optimization: Bulk Selection)
    sel.Clear
    Dim i As Long
    For i = 1 To children.count
        sel.Add children.item(i)
    Next
    visp.SetShow 1 ' Hide
    sel.Clear
    ' Iterate and Capture
    For i = 1 To children.count
        Dim child As Product: Set child = children.item(i)
        ' Show current child
        sel.Add child
        visp.SetShow 0 ' Show
        sel.Clear
        ' Recurse
        CaptureRecursive child, folderPath
        sel.Add child
        visp.SetShow 1 ' Hide
        sel.Clear
    Next
    For i = 1 To children.count
        sel.Add children.item(i)
    Next
    visp.SetShow 0
    sel.Clear
End Sub

Private Sub InitializeSettings()
    With m_Settings
        .BackgroundColor = Array(1, 1, 1) ' White
        .FocusFactor = DEFAULT_FOCUS_FACTOR
        .ViewDirection = Array(-1, -1, -1) ' Isometric
        .RenderingMode = 1 ' Shading with Edges
        .AspectHeightRatio = 0.618
        .ResolutionWidth = 1080
    End With
    m_FirstImagePath = ""
End Sub

Private Sub SetCAPDisplay()
    With CATIA.ActiveWindow
        .WindowState = 0 ' Maximized
        .Width = m_Settings.ResolutionWidth
        .Height = .Width * m_Settings.AspectHeightRatio
        .Layout = 1 ' Geometry only
    End With
    Dim oViewer: Set oViewer = CATIA.ActiveWindow.ActiveViewer
    With oViewer
        .RenderingMode = m_Settings.RenderingMode
        .Viewpoint3D.PutSightDirection m_Settings.ViewDirection
        .Reframe
        .Viewpoint3D.FocusDistance = .Viewpoint3D.FocusDistance * m_Settings.FocusFactor
        .PutBackgroundColor m_Settings.BackgroundColor
    End With
End Sub
Private Sub RecoverDisplay()
    CATIA.ActiveWindow.Layout = 2 ' Specs and Geometry
    Dim oViewer: Set oViewer = CATIA.ActiveWindow.ActiveViewer
    oViewer.PutBackgroundColor Array(0.2, 0.2, 0.4) ' Default Dark Blue
    oViewer.Reframe
End Sub

Private Sub HideNonBody(iDoc)
    Dim sel As Selection: Set sel = iDoc.Selection: sel.Clear
    Dim searchStr As String
    searchStr = ".Plane + .AxisSystem + .Point + .2DPoint + .Curve + .2DCurve + .Surface + .MfConstraint,all"
    
''    You may have noticed the ,all or ,sel. There are multiple ways to search:
'
'Everywhere: shortcut “all”
'InWorkbench: shortcut “in”
'FromWorkbench: shortcut “from”
'FromSelection: shortcut “sel”
'VisibleOnScreen: shortcut “scr”
    
    
    On Error Resume Next
    sel.Search searchStr
    If sel.count > 0 Then
        sel.VisProperties.SetShow 1 ' 隐藏 (catVisPropertyNoShow)
        sel.Clear
    End If
    On Error GoTo 0
End Sub



