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
Const CF_BITMAP = 2
Private Const Fdis = 0.9
Private thisdir

Sub Capturetopath()

If Not KCL.CanExecute("ProductDocument") Then Exit Sub
    On Error Resume Next
     CATIA.StartCommand ("* iso")
      Dim btn, bTitle, bResult
      imsg = "如要截图，请等待ISO视角调整完毕后点击确认"
        btn = vbYesNo + vbExclamation
        bResult = MsgBox(imsg, btn, "bTitle")  ' Yes(6),No(7),cancel(2)
        Select Case bResult
            Case 7: Exit Sub '===选择“否”====
            Case 2: Exit Sub '===选择“取消”====
            Case 6  '===选择“是”====
                Call Capme
            End Select
  
  If Err.Number = 0 Then
     KCL.openpath (thisdir)
   End If
     Err.Clear
On Error GoTo 0
  

End Sub

Sub Capme()
 If Not KCL.CanExecute("ProductDocument,PartDocument") Then Exit Sub
 If pdm Is Nothing Then
        Set pdm = New class_PDM
 End If
On Error Resume Next
'-----------设置显示样式模式-------------
 Call HideNonBody(rootDoc, 1)
    CATIA.RefreshDisplay = True
    CATIA.DisplayFileAlerts = False
 With CATIA.Application
   .Width = 1920 / 2
   .Height = 1080 '.Width * 0.618
 End With
    
With CATIA.ActiveWindow
     .WindowState = 0  '   '0 catWindowStateMaximized 1   catWindowStateMinimized,2   catWindowStateNormal
     .Width = 1080
     .Height = .Width * 0.618
     .Layout = 1    ' 仅显示几何视图
End With

  CATIA.RefreshDisplay = False
     Dim oViewer
     Set oViewer = CATIA.ActiveWindow.ActiveViewer
     With oViewer
        .RenderingMode = 1 ' catRenderShadingWithEdges
        .Viewpoint3D.PutSightDirection Array(-1, -1, -1)
        .Reframe
        .Viewpoint3D.FocusDistance = oViewer.Viewpoint3D.FocusDistance * Fdis
        .PutBackgroundColor Array(1, 1, 1) '白色背景
     End With
         
    CATIA.StartCommand ("Compass")  '隐藏指南针
     Dim oprd
     Set oprd = rootprd
     If oprd Is Nothing Then Exit Sub
     oprd.ApplyWorkMode (3)  '3  DESIGN_MODE
     Dim oPath
     oPath = KCL.GetPath(KCL.getVbaDir & "\" & "oTemp")
     KCL.ClearDir (oPath) '截图前先清空文件夹
     
     
     If gPic_Path = "" Then
            gPic_Path = oPath
     End If
     allPN.Remove all
     
     CaptureMe oprd, oPath
'-----------恢复显示样式模式-------------
     CATIA.DisplayFileAlerts = True
     owd.WindowState = 0
     oViewer.PutBackgroundColor Array(0.2, 0.2, 0.4)
     CATIA.RefreshDisplay = True
     CATIA.ActiveWindow.Layout = 2 ' catWindowSpecsAndGeom
     CATIA.StartCommand ("Compass")
     allPN.Remove all
     Set oprd = Nothing

On Error GoTo 0

End Sub
Sub CaptureMe(iprd, oFolder)
    On Error Resume Next
     Dim oViewer
     Set oViewer = CATIA.ActiveWindow.ActiveViewer
     oViewer.RenderingMode = 1 ' catRenderShadingWithEdges
     oViewer.Viewpoint3D.PutSightDirection Array(-1, -1, -1)
     oViewer.Reframe
     oViewer.Viewpoint3D.FocusDistance = oViewer.Viewpoint3D.FocusDistance * Fdis
    
     If allPN.Exists(iprd.PartNumber) = False Then  '对产品截图并遍历
       allPN(iprd.PartNumber) = 1
         imgfilename = oFolder & "\" & iprd.ReferenceProduct.PartNumber & ".jpg"
          oViewer.CaptureToFile 5, imgfilename
     End If
     If thisdir = "" Then
          thisdir = imgfilename
     End If
          
    Dim osel: Set osel = CATIA.ActiveDocument.Selection
    osel.Clear
    Dim VisPoSel: Set VisPoSel = osel.VisProperties
    Dim children, i
    Set children = iprd.Products
    
    For Each cPrd In children
    osel.Add cPrd
    Next
    VisPoSel.SetShow 1
    osel.Clear      ' 隐藏所有子产品
    If children.count > 0 Then
                         For i = 1 To children.count     ' 递归处理每个子产品
                              osel.Add children.item(i)
                              VisPoSel.SetShow 0
                              osel.Clear '显示当前子产品
                        
                        Call CaptureMe(children.item(i), oFolder)
                        
                        osel.Add children.item(i) ' 隐藏当前子产品
                        VisPoSel.SetShow 1
                        osel.Clear
                    Next
   End If
   
     For Each cPrd In children ' 重新显示每个子产品
       osel.Add cPrd
     Next
       VisPoSel.SetShow 0
       osel.Clear
       
End Sub
Sub HideNonBody(iDoc, catVisPropertyNoShowAttr As Integer)
     On Error Resume Next
     Dim osel As Selection
     Set osel = iDoc.Selection
     osel.Clear
     osel.Search "(((CATStFreeStyleSearch.Plane + CATPrtSearch.Plane) + CATGmoSearch.Plane) + CATSpdSearch.Plane),all"
     osel.VisProperties.SetShow catVisPropertyNoShowAttr
     osel.Clear
     osel.Search "(((CATStFreeStyleSearch.AxisSystem + CATPrtSearch.AxisSystem) + CATGmoSearch.AxisSystem) + CATSpdSearch.AxisSystem),all"
     osel.VisProperties.SetShow catVisPropertyNoShowAttr
     osel.Clear
     osel.Search "((((((CATStFreeStyleSearch.Point + CAT2DLSearch.2DPoint) + CATSketchSearch.2DPoint) + CATDrwSearch.2DPoint) + CATPrtSearch.Point) + CATGmoSearch.Point) + CATSpdSearch.Point),all"
     osel.VisProperties.SetShow catVisPropertyNoShowAttr
     osel.Clear
     osel.Search "((((((CATStFreeStyleSearch.Curve + CAT2DLSearch.2DCurve) + CATSketchSearch.2DCurve) + CATDrwSearch.2DCurve) + CATPrtSearch.Curve) + CATGmoSearch.Curve) + CATSpdSearch.Curve),all"
     osel.VisProperties.SetShow catVisPropertyNoShowAttr
     osel.Clear
     osel.Search "(((CATStFreeStyleSearch.Surface + CATPrtSearch.Surface) + CATGmoSearch.Surface) + CATSpdSearch.Surface),all"
     osel.VisProperties.SetShow catVisPropertyNoShowAttr
     osel.Clear
     osel.Search "(((((((CATProductSearch.MfConstraint + CATStFreeStyleSearch.MfConstraint) + CATAsmSearch.MfConstraint) + CAT2DLSearch.MfConstraint) + CATSketchSearch.MfConstraint) + CATDrwSearch.MfConstraint) + CATPrtSearch.MfConstraint) + CATSpdSearch.MfConstraint),all"
     osel.VisProperties.SetShow catVisPropertyNoShowAttr
     osel.Clear
     Err.Clear
     On Error GoTo 0
End Sub
