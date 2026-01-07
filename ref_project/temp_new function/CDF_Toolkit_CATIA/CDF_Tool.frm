VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} CDF_Tool 
   Caption         =   "臭豆腐工具箱CATIA版"
   ClientHeight    =   9690.001
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   2340
   OleObjectBlob   =   "CDF_Tool.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "CDF_Tool"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub BatchRename_Click()
Batch_Rename_3D.CATMain
End Sub

Private Sub cmdActRenameNumber_Click()
CNC_ReNumberAct.CATMain
End Sub

Private Sub cmdAutoDrw_Click()
AutoDrw_2D_3D.CATMain
Unload CDF_Tool
End Sub

Private Sub cmdBatchPrint_Click()
Batch_Print_2D3D.CATMain
End Sub

Private Sub cmdBatchPrt_Click()
Batch_Print_2D3D.CATMain
End Sub

Private Sub cmdCaptureImage_Click()
Capture_Image_3D.CATMain
End Sub

Private Sub cmdCNCTools_Click()
CNC_Prog_List2.CATMain
End Sub

Private Sub cmdColorShow_Click()
Body_Color_3D.CATMain
End Sub

Private Sub cmdDimensionTag_Click()
Dimension_Tag_2D.CATMain
End Sub

Private Sub cmdDrwAuto_Click()
AutoDrw_2D_3D.CATMain
Unload CDF_Tool
End Sub

Private Sub cmdDrwLink_Click()
Change_Link_2D.CATMain
End Sub

Private Sub cmdExportPDF_Click()
IntCATIA
ExportPDF oActDoc
End Sub

Private Sub cmdExportSTP2_Click()
Export_Stp_2D3D.CATMain
End Sub

Private Sub cmdExportTable_Click()
Table_2_Excel_2D.CATMain
End Sub

Private Sub cmdFix_Click()
Fix_Unfix_Position_3D.CATMain
End Sub

Private Sub cmdGBstyle_Click()
View_Name_GB_2D.CATMain
End Sub

Private Sub cmdGraphReOrder_Click()
Graph_ReOrder_3D.CATMain
End Sub

Private Sub cmdGrid_Click()
Draw_Cor_Grid_2D.CATMain
End Sub

Private Sub cmdInsertDrw_Click()
Insert_GB_Frame_3D.CATMain
End Sub

Private Sub cmdNewfrmTemplate_Click()
New_From_Template.CATMain
End Sub

Private Sub cmdOpenCurrentFd_Click()
Open_Current_Folder.CATMain
End Sub

Private Sub cmdOpenCurrentFolder_Click()
Open_Current_Folder.CATMain
End Sub

Private Sub cmdOpenModelDrafting_Click()
Open_Drawing_Model_3D.CATMain
End Sub

Private Sub cmdOpenModelDwg_Click()
Open_Drawing_Model_3D.CATMain
End Sub

Private Sub cmdPjPlane_Click()
Modify_Proj_Plane_2D.CATMain
End Sub

Private Sub cmdPrgList_Click()
CNC_Prog_List2.CATMain
End Sub

Private Sub cmdPrgListExcel_Click()
CNC_Prog_List.CATMain
End Sub

Private Sub cmdPrjPlane_Click()
Modify_Proj_Plane_2D.CATMain
End Sub



Private Sub cmdReg_Click()
On Error Resume Next
Shell oCATVBA_Folder.Path & "\" & "Reg.exe", vbNormalFocus
If Err.Number <> 0 Then
MsgBox "无法调用Reg.exe！" & vbCrLf & "请确认Reg.exe存在于" & oCATVBA_Folder.Path & vbCrLf & "你也可以找到此Reg.exe双击运行注册。", vbQuestion, "臭豆腐工具箱CATIA版"
End If
End Sub

Private Sub cmdTable_Click()
Table_2_Excel_2D.CATMain
End Sub

Private Sub cmdWrkFolder_Click()
Open_Current_Folder.CATMain
End Sub

Private Sub cmdZipPDFSTP_Click()
Export_Zipped_PDF_STP_2D.CATMain
End Sub

Private Sub CommandButton2_Click()
Open_Current_Folder.CATMain
End Sub

Private Sub Export_BOM2_Click()
Exact_BOM_3D.CATMain
Unload CDF_Tool
End Sub

Private Sub MultiPage1_Change()
Select Case MultiPage1.Value
Case 0
MultiPage1.Height = 486
Case 1
MultiPage1.Height = 405
Case 2
MultiPage1.Height = 256
Case 3
MultiPage1.Height = 486
Case Else
End Select
Me.Height = MultiPage1.Height + 26
Me.Width = MultiPage1.Width + 10
End Sub

Private Sub NC_Prog_ReName_Click()
MP_ReName_Export_3D.CATMain
End Sub

Private Sub Properties_Click()
Add_Properties_3D.CATMain
End Sub

Private Sub Export_PDF_Click()
IntCATIA
ExportPDF oActDoc
End Sub

Private Sub Export_STP_Click()
IntCATIA
ExportStp oActDoc
End Sub

Private Sub Inport_GB_Frame_Click()
Insert_GB_Frame_3D.CATMain
'Unload CDF_Tool
End Sub

Private Sub Part_2_Product_Click()
Part_2_Product_3D.CATMain
End Sub


'*******************************************************************
'*******************************************************************

Sub ExportPDF(oDocPDF, Optional ByVal sPath As String = "") 'sPath为输出目录，不带“\”
On Error Resume Next
Dim sName
If oDocPDF.Saved = False Then
   If MsgBox("要保存修改么?", vbYesNo) = vbYes Then
    oDocPDF.Save
   End If
End If

If (TypeName(oDocPDF) <> "DrawingDocument") Then
    MsgBox "此命令只能在工程制图模块中运行", vbInformation, "Information"
    Exit Sub
End If
Dim settingControllers1 'As SettingControllers
Dim settingRepository1 'As SettingRepository
Dim tempB
Set settingControllers1 = CATIA.SettingControllers
Set settingRepository1 = settingControllers1.Item("DraftingOptions")
tempB = settingRepository1.GetAttr("DimDesignMode")
'msgbox tempB
settingRepository1.PutAttr "DimDesignMode", False
'msgbox settingRepository1.GetAttr("DimDesignMode")
settingRepository1.Commit
If sPath = "" Then
    sName = oDocPDF.Path & "\" & Replace(oDocPDF.Name, ".", "_")
Else
   sName = sPath & "\" & Replace(oDocPDF.Name, ".", "_")
End If
    oDocPDF.ExportData sName, "pdf"
    
settingRepository1.PutAttr "DimDesignMode", tempB
settingRepository1.Commit
On Error GoTo 0
End Sub

Sub ExportStp(oDocStp, Optional ByVal sPath As String = "") 'sPath为输出目录，不带“\”
On Error Resume Next
Dim sName
If oDocStp.Saved = False Then
   If MsgBox("要保存修改么?", vbYesNo) = vbYes Then
    oDocStp.Save
    End If
End If
'*****************************
Dim settingControllers1 As Object
Dim stepSettingAtt1 As Object
Dim short1

Set settingControllers1 = CATIA.SettingControllers
Set stepSettingAtt1 = settingControllers1.Item("CATSdeStepSettingCtrl")
short1 = stepSettingAtt1.AttAP
'Debug.Print stepSettingAtt1.AttAP
If short1 <> 2 Then
stepSettingAtt1.AttAP = 2
'Debug.Print stepSettingAtt1.AttAP
End If

If (TypeName(oDocStp) = "PartDocument") Or (TypeName(oDocStp) = "ProductDocument") Then
    If sPath = "" Then
    sName = oDocStp.Path & "\" & Replace(oDocStp.Name, ".", "_")
    Else
    sName = sPath & "\" & Replace(oDocStp.Name, ".", "_")
    End If
    oDocStp.ExportData sName, "stp"
End If

If (TypeName(oDocStp) = "DrawingDocument") Then
    Dim oDocStp1
    Set oDocStp1 = DwgLinkedDoc(oDocStp)
    If sPath = "" Then
    sName = oDocStp1.Path & "\" & Replace(oDocStp1.Name, ".", "_")
    Else
    sName = sPath & "\" & Replace(oDocStp1.Name, ".", "_")
    End If
    oDocStp1.ExportData sName, "stp"
End If
stepSettingAtt1.AttAP = short1
'Debug.Print stepSettingAtt1.AttAP
Set settingControllers1 = Nothing
Set stepSettingAtt1 = Nothing
On Error GoTo 0
End Sub






'View must parallel to system aixes,view angle 0deg,90deg and -90deg

Sub Draw_Grid(oDocDrw)

If TypeName(oDocDrw) <> "DrawingDocument" Then
    MsgBox "此命令只能在工程制图模块下运行!"
    Exit Sub
End If


CATIA.RefreshDisplay = False
Dim sStatus As String

    ' Set the CATIA popup file alerts to False
    ' It prevents to stop the macro at each alert during its execution
CATIA.DisplayFileAlerts = False

    ' Optional: allows to find the sample wherever it's installed
  
    ' Variables declaration
    Dim oDrwDocument As DrawingDocument
    Set oDrwDocument = oDocDrw
    
    Dim oDrwSheets As DrawingSheets
    Dim oDrwSheet As DrawingSheet
    Dim oDrwView As DrawingView
    Dim oFactory2D ' As Factory2D
   
   ' The Distance between the lines
    Dim d As Integer
    Dim nx As Integer
    Dim ny As Integer

   ' The point coordinate select from Drawing
    Dim x1 As Integer
    Dim y1 As Integer
    Dim x2 As Integer
    Dim y2 As Integer
    Dim Pt1 'As Point2D
    Dim Pt2 'As Point2D
    
   'The view scale dAngle for rotate view scale for view scale
    Dim dScale, dAngle As Double
  
    'The view coordinate origin
    Dim x, y As Integer
    On Error Resume Next
    d = InputBox("准备工作:" & vbCrLf & "请激活需要添加百格线的视图并预先定义两点以确定百格线绘制范围" & vbCrLf & "点1为绘制区左下角" & vbCrLf & "点2为绘制区右上角" & vbCrLf & vbCrLf & "完成准备工作后再输入间距否则请先关闭此对话框" & vbCrLf & vbCrLf & "请输入间距：", "input box", "100")
    d = CInt(d)
    If Err.Number <> 0 Then
    Exit Sub
    End If
       

    Set oDrwSheets = oDrwDocument.Sheets
    Set oDrwSheet = oDrwSheets.ActiveSheet
    Set oDrwView = oDrwSheet.Views.ActiveView
    Set oFactory2D = oDrwView.Factory2D
  
    'Retrive the value of the view
     x = oDrwView.xAxisData
     y = oDrwView.yAxisData
     dScale = oDrwView.Scale
     dAngle = oDrwView.Angle
     

   'Get the coordinate from the select two point
    'On Error Resume Next

     Set oSel = oDocDrw.Selection
     oSel.Clear

     ReDim sFilter(0)
     sFilter(0) = "Point2D"
     MsgBox "请选择点 1 以定义绘制区左下角 "
     sStatus = oSel.SelectElement2(sFilter, "请选择点 1 以定义绘制区左下角", False)
     If (sStatus = "Normal") Then

      Dim SelectedPoint1
      Set SelectedPoint1 = oSel.Item(1)

      Dim pt1Coord(2)
      SelectedPoint1.GetCoordinates pt1Coord
      'MsgBox "The frst point has been selected "

      Else: MsgBox "你未选择绘制区顶点,程序退出"
      Exit Sub
      End If

      MsgBox "请选择点 2 以定义绘制区右上角 "
      sStatus = oSel.SelectElement2(sFilter, "请选择点 2 以定义绘制区右上角", False)
      If (sStatus = "Normal") Then

       Dim SelectedPoint2

       Set SelectedPoint2 = oSel.Item(1)

      Dim pt2Coord(2)
    SelectedPoint2.GetCoordinates pt2Coord
     'MsgBox "The second point has been selected "

     Else: MsgBox "你未选择完成绘制区顶点,程序退出"
     Exit Sub
     End If
   
  If dAngle = 0 Then
   x1 = CInt((pt1Coord(0) - x) / dScale)
   y1 = CInt((pt1Coord(1) - y) / dScale)
   x2 = CInt((pt2Coord(0) - x) / dScale)
   y2 = CInt((pt2Coord(1) - y) / dScale)
  End If
   If dAngle > 0 Then
     x1 = CInt((pt1Coord(1) - y) / dScale)
    y1 = CInt((pt1Coord(0) - x) / dScale)
    x2 = CInt((pt2Coord(1) - y) / dScale)
    y2 = CInt((pt2Coord(0) - x) / dScale)
  End If
 
  If dAngle < 0 Then
     x1 = CInt((pt1Coord(1) - y) / dScale)
    y1 = CInt((pt1Coord(0) - x) / dScale)
    x2 = CInt((pt2Coord(1) - y) / dScale)
    y2 = CInt((pt2Coord(0) - x) / dScale)
  End If

    x1 = d * CInt(x1 / d)
    y1 = d * CInt(y1 / d)
    x2 = d * CInt(x2 / d)
    y2 = d * CInt(y2 / d)
    
    nx = (x2 - x1) \ d 'The number of the horizontal line
    ny = (y2 - y1) \ d 'The number of the vertical  line
    

    Dim Line2D1 ' As Line2D
    Dim Circle2D1 'As Circle2D
   Dim MyText As DrawingText
   Dim iFontSize As Double
   
    Dim i, j, R
   
    iFontSize = 3.5
    R = 6
    R = R / dScale

'------------------------------------------------------
Dim Di_H, Di_V
Dim Text_XYZ_H As String
Dim Text_XYZ_V As String

Di_H = 1
Di_V = 1

'Compare the drawing view HV with 3D Aixes
'Dim XX1, YY1, ZZ1, XX2, YY2, ZZ2 'As Integer
Dim XX1 As Double
Dim YY1 As Double
Dim ZZ1 As Double
Dim XX2 As Double
Dim YY2 As Double
Dim ZZ2 As Double

oDrwView.GenerativeBehavior.GetProjectionPlane XX1, YY1, ZZ1, XX2, YY2, ZZ2
MsgBox "XX1=" & XX1 & "  /YY1=" & YY1 & "  /ZZ1=" & ZZ1 & vbCrLf & "XX2=" & XX2 & "  /YY2=" & YY2 & "  /ZZ2=" & ZZ2


If (XX1 = 1) Then
    Text_XYZ_H = "X"
End If
If (XX1 = -1) Then
    Text_XYZ_H = "X"
    Di_H = -1
End If

If (YY1 = 1) Then
    Text_XYZ_H = "Y"
End If
If (YY1 = -1) Then
    Text_XYZ_H = "Y"
    Di_H = -1
End If

If (ZZ1 = 1) Then
    Text_XYZ_H = "Z"
End If
If (ZZ1 = -1) Then
    Text_XYZ_H = "Z"
    Di_H = -1
End If


If (XX2 = 1) Then
    Text_XYZ_V = "X"
End If
If (XX2 = -1) Then
    Text_XYZ_V = "X"
    Di_V = -1
End If

If (YY2 = 1) Then
    Text_XYZ_V = "Y"
End If
If (YY2 = -1) Then
    Text_XYZ_V = "Y"
    Di_V = -1
End If

If (ZZ2 = 1) Then
    Text_XYZ_V = "Z"
End If
If (ZZ2 = -1) Then
    Text_XYZ_V = "Z"
    Di_V = -1
End If

If dAngle > 0 Then
    Di_V = -Di_V
End If
If dAngle < 0 Then
    Di_H = -Di_H
End If

Dim oVisProps As VisPropertyset
oSel.Clear

Dim TextV
TextV = R / 2

 'Draw the  horizontall line

    For i = 0 To ny
      If dAngle = 0 Then
          Set Line2D1 = oFactory2D.CreateLine(x1 - d / 3, y1 + d * i, x1 + nx * d + d / 3, y1 + d * i)
          oSel.Add Line2D1
          Set Circle2D1 = oFactory2D.CreateClosedCircle(x1 - d / 3 - R, y1 + d * i, R)
          oSel.Add Circle2D1
          Set Line2D1 = oFactory2D.CreateLine(x1 - d / 3 - R * 2, y1 + d * i, x1 - d / 3, y1 + d * i)
          oSel.Add Line2D1
          Set MyText = oDrwView.Texts.Add(Text_XYZ_V, x1 - d / 3 - R, y1 + d * i + TextV)
         MyText.AnchorPosition = catMiddleCenter
         MyText.SetFontSize 0, 0, iFontSize
          Set MyText = oDrwView.Texts.Add((y1 + d * i) * Di_V, x1 - d / 3 - R, y1 + d * i - TextV)
          MyText.AnchorPosition = catMiddleCenter
          MyText.SetFontSize 0, 0, iFontSize
       End If

      If dAngle > 0 Then
          Set Line2D1 = oFactory2D.CreateLine(x1 - d / 3, -y1 - d * i, x1 + nx * d + d / 3, -y1 - d * i)
          oSel.Add Line2D1
          Set Circle2D1 = oFactory2D.CreateClosedCircle(x1 + nx * d + d / 3 + R, -y1 - d * i, R)
          oSel.Add Circle2D1
          Set Line2D1 = oFactory2D.CreateLine(x1 + nx * d + d / 3 + R, -y1 - d * i + R, x1 + nx * d + d / 3 + R, -y1 - d * i - R)
          oSel.Add Line2D1
          Set MyText = oDrwView.Texts.Add(Text_XYZ_V, x1 + nx * d + d / 3 + R + TextV, -y1 - d * i)
          MyText.AnchorPosition = catMiddleCenter
          MyText.SetFontSize 0, 0, iFontSize
          Set MyText = oDrwView.Texts.Add((y1 + d * i) * Di_V, x1 + nx * d + d / 3 + R - TextV, -y1 - d * i)
          MyText.AnchorPosition = catMiddleCenter
          MyText.SetFontSize 0, 0, iFontSize
       End If

      If dAngle < 0 Then
          Set Line2D1 = oFactory2D.CreateLine(-x1 + d / 3, y1 + d * i, -(x1 + nx * d + d / 3), y1 + d * i)
          oSel.Add Line2D1
          Set Circle2D1 = oFactory2D.CreateClosedCircle(-(x1 + nx * d + d / 3) - R, y1 + d * i, R)
          oSel.Add Circle2D1
          Set Line2D1 = oFactory2D.CreateLine(-x1 - nx * d - d / 3 - R, y1 + d * i + R, -x1 - nx * d - d / 3 - R, y1 + d * i - R)
          oSel.Add Line2D1
          Set MyText = oDrwView.Texts.Add(Text_XYZ_V, -x1 - nx * d - d / 3 - R + TextV, y1 + d * i)
          MyText.AnchorPosition = catMiddleCenter
          MyText.SetFontSize 0, 0, iFontSize
          Set MyText = oDrwView.Texts.Add((y1 + d * i) * Di_V, -x1 - nx * d - d / 3 - R - TextV, y1 + d * i)
          MyText.AnchorPosition = catMiddleCenter
         MyText.SetFontSize 0, 0, iFontSize
       End If

    Next


    
  'Draw the vertical  line
    For j = 0 To nx
      If dAngle = 0 Then
          Set Line2D1 = oFactory2D.CreateLine(x1 + d * j, y1 - d / 3, x1 + d * j, y1 + ny * d + d / 3)
          oSel.Add Line2D1
          Set Circle2D1 = oFactory2D.CreateClosedCircle(x1 + d * j, y1 + ny * d + d / 3 + R, R)
          oSel.Add Circle2D1
          Set Line2D1 = oFactory2D.CreateLine(x1 + d * j - R, y1 + ny * d + d / 3 + R, x1 + d * j + R, y1 + ny * d + d / 3 + R)
          oSel.Add Line2D1
           Set MyText = oDrwView.Texts.Add(Text_XYZ_H, x1 + d * j, y1 + ny * d + d / 3 + R + TextV)
                 MyText.AnchorPosition = catMiddleCenter
                 MyText.SetFontSize 0, 0, iFontSize
           Set MyText = oDrwView.Texts.Add((x1 + d * j) * Di_H, x1 + d * j, y1 + ny * d + d / 3 + R - TextV)
                 MyText.AnchorPosition = catMiddleCenter
                 MyText.SetFontSize 0, 0, iFontSize
      End If

      If dAngle > 0 Then
          Set Line2D1 = oFactory2D.CreateLine(x1 + d * j, -y1 + d / 3, x1 + d * j, -y1 - ny * d - d / 3)
          oSel.Add Line2D1
          Set Circle2D1 = oFactory2D.CreateClosedCircle(x1 + d * j, -y1 + d / 3 + R, R)
          oSel.Add Circle2D1
          Set Line2D1 = oFactory2D.CreateLine(x1 + d * j, -y1 + d / 3 + R * 2, x1 + d * j, -y1 + d / 3)
          oSel.Add Line2D1
           Set MyText = oDrwView.Texts.Add(Text_XYZ_H, x1 + d * j + TextV, -y1 + d / 3 + R)
                 MyText.AnchorPosition = catMiddleCenter
                 MyText.SetFontSize 0, 0, iFontSize
           Set MyText = oDrwView.Texts.Add((x1 + d * j) * Di_H, x1 + d * j - TextV, -y1 + d / 3 + R)
                 MyText.AnchorPosition = catMiddleCenter
                 MyText.SetFontSize 0, 0, iFontSize
      End If

      If dAngle < 0 Then
          Set Line2D1 = oFactory2D.CreateLine(-x1 - d * j, y1 - d / 3, -x1 - d * j, y1 + ny * d + d / 3)
          oSel.Add Line2D1
          Set Circle2D1 = oFactory2D.CreateClosedCircle(-x1 - d * j, y1 - d / 3 - R, R)
          oSel.Add Circle2D1
          Set Line2D1 = oFactory2D.CreateLine(-x1 - d * j, y1 - d / 3 - R * 2, -x1 - d * j, y1 - d / 3)
          oSel.Add Line2D1
           Set MyText = oDrwView.Texts.Add(Text_XYZ_H, -x1 - d * j + TextV, y1 - d / 3 - R)
                 MyText.AnchorPosition = catMiddleCenter
                 MyText.SetFontSize 0, 0, iFontSize
           Set MyText = oDrwView.Texts.Add((x1 + d * j) * Di_H, -x1 - d * j - TextV, y1 - d / 3 - R)
                 MyText.AnchorPosition = catMiddleCenter
                 MyText.SetFontSize 0, 0, iFontSize
      End If

   Next
Dim oFontSize As Long
  ' MyText.SetFontSize 0,  0, iFontSize
    Set oVisProps = oSel.VisProperties
    oVisProps.SetRealWidth 1, 0 '1st parameter line width 1-63 2nd parameter inheritance flag 1 or 0
    oVisProps.SetRealColor 0, 255, 0, 1

    Set oVisProps = Nothing
    Set oSel = Nothing
   


    ' Update drawing table modifications

    CATIA.ActiveWindow.ActiveViewer.Reframe
    
End Sub


Private Sub Send2SubFolder_Click()
Send_2_SubFolder_3D.CATMain
End Sub

Private Sub UserForm_Activate()
MultiPage1.Value = 0
'On Error Resume Next
Select Case CATIA.GetWorkbenchId()
Case "Drw"
    MultiPage1.Value = 1
Case "LatheProgramWorkbench"
    MultiPage1.Value = 2
Case "ManufacturingProgramWorkbench"
    MultiPage1.Value = 2
Case "M3xProgramWorkbench"
    MultiPage1.Value = 2
Case "CATAMGProgramWorkbench"
    MultiPage1.Value = 2
Case "ManufacturingNCReviewWorkbench"
    MultiPage1.Value = 2
Case Else
End Select
If oActDoc Is Nothing Then
MultiPage1.Value = 0
End If
Dim Picfile
Picfile = "Chtang80.jpg"
Picfile = oCATVBA_Folder("Config").Path & "\" & Picfile
'Image1.Picture = LoadPicture(Picfile)

End Sub
