VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Draw_Grid 
   Caption         =   "臭豆腐工具箱|百格线"
   ClientHeight    =   4830
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   2310
   OleObjectBlob   =   "Draw_Grid.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "Draw_Grid"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
#If VBA7 Then
Private Declare PtrSafe Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As LongPtr
Private Declare PtrSafe Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As LongPtr
#Else
Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Private Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long
#End If


Private Sub cmdDelGrid_Click()
IntCATIA
If TypeName(oActDoc) <> "DrawingDocument" Then
    MsgBox "此命令只能在工程制图模块下运行!", vbOKOnly, "百格线工具"
    Exit Sub
End If

Dim oDrwView As Object 'DrawingView

CATIA.RefreshDisplay = False
CATIA.DisplayFileAlerts = False

Set oDrwView = oActDoc.Sheets.ActiveSheet.Views.ActiveView
On Error Resume Next
    If oDrwView.ViewType = 13 Or oDrwView.ViewType = 0 Then
       If MsgBox("请选择包含百格线的视图", vbOKCancel, "百格线工具") = vbOK Then
          Set oDrwView = MProg.Sel("DrawingView")
          oDrwView.Activate
       Else
          Exit Sub
       End If
    End If
    If Err.Number <> 0 Then
    Exit Sub
    End If
 'Programmed by Charles.Tang 13507468076
On Error GoTo 0

Set oSel = oActDoc.Selection
oSel.Clear

Dim oDrwText As Object
Dim Line2D1 As Object

For Each oDrwText In oDrwView.Texts
    If Left(oDrwText.Name, 5) = "Grid_" Then
    oSel.Add oDrwText
    End If
Next

For Each Line2D1 In oDrwView.GeometricElements
    If Line2D1.GeometricType = 3 And Left(Line2D1.Name, 5) = "Grid_" Then
    oSel.Add Line2D1
    End If
Next
If oSel.Count = 0 Then
   If MsgBox("当前激活视图不包含臭豆腐工具箱生成的百格线" & _
        vbCrLf & "需要选择另一个视图吗?" & _
        vbCrLf & "确定[OK]选择另一个视图, 取消[Cancel]退出", vbOKCancel, "百格线工具") = vbCancel Then
        Exit Sub
    Else
            For Each oDrwView In oActDoc.Sheets.ActiveSheet.Views
            If oDrwView.ViewType = 13 Then
               oDrwView.Activate
            End If
            Next
        Call cmdDelGrid_Click
    End If
Else
    '如果视图被锁定则解除锁定
    'MsgBox oDrwView.LockStatus
    If oDrwView.LockStatus = 1 Then
       oDrwView.LockStatus = 0
       oSel.Delete
       oDrwView.LockStatus = 1
    Else
       oSel.Delete
    End If

End If

'将激活视图设为主视图
For Each oDrwView In oActDoc.Sheets.ActiveSheet.Views
    If oDrwView.ViewType = 13 Then
       oDrwView.Activate
    End If
Next


'恢复命令执行前的设置
CATIA.RefreshDisplay = True
CATIA.DisplayFileAlerts = True

Set oDrwView = Nothing
Set oSel = Nothing
Set Line2D1 = Nothing
Set oDrwText = Nothing
End Sub

Private Sub cmdRunGrid_Click()
IntCATIA
If TypeName(oActDoc) <> "DrawingDocument" Then
    MsgBox "此命令只能在工程制图模块下运行!", vbOKOnly, "百格线工具"
    Exit Sub
End If


'Dim oDrwSheets As Object ' DrawingSheets
'Dim oDrwSheet As Object 'DrawingSheet
'Dim oDrwViews As Object 'DrawingViews
Dim oDrwView As Object 'DrawingView



CATIA.RefreshDisplay = False
CATIA.DisplayFileAlerts = False


    For Each oDrwView In oActDoc.Sheets.ActiveSheet.Views
    If oDrwView.ViewType = 13 Then
       oDrwView.Activate
    End If
    Next
On Error Resume Next
Set oDrwView = MProg.Sel("DrawingView")
If Err.Number <> 0 Then
Exit Sub
End If

oDrwView.Activate

'MsgBox oDrwView.Name
'检查投影方向是否垂直于XYZ平面，如果不是则退出
Dim Xn As Double
Dim Yn As Double
Dim Zn As Double

Dim x1 As Double
Dim y1 As Double
Dim Z1 As Double
Dim x2 As Double
Dim y2 As Double
Dim Z2 As Double
'
oDrwView.GenerativeBehavior.GetProjectionPlaneNormal Xn, Yn, Zn
'MsgBox "ProjectNormal is (" & Xn & "," & Yn & "," & Zn & ")"
oDrwView.GenerativeBehavior.GetProjectionPlane x1, y1, Z1, x2, y2, Z2
'MsgBox "ProjectPlane is (" & X1 & "," & Y1 & "," & Z1 & ")," & "(" & X2 & "," & Y2 & "," & Z2 & ")"

Dim scheck As String
If Not (Abs(Xn) = 1 Or Abs(Yn) = 1 Or Abs(Zn) = 1) Then
    scheck = "该视图投影方向是 (" & Round(Xn, 2) & "," & Round(Yn, 2) & "," & Round(Zn, 2) & "), 并不与基准坐标平面垂直！" & _
    vbCrLf & "该视图不能生成正确的百格线,程序退出！"
ElseIf Not (Abs(x1) = 1 Or Abs(y1) = 1 Or Abs(Z1) = 1 Or Abs(x2) = 1 Or Abs(y2) = 1 Or Abs(Z2) = 1) Then
        scheck = scheck & vbCrLf & "该视图投影平面的XY轴与基准坐标系的坐标轴不平行！" & _
        vbCrLf & "投影平面的XY轴在基准坐标系中的方向是 " & _
        vbCrLf & "(" & Round(x1, 2) & "," & Round(y1, 2) & "," & Round(Z1, 2) & ")," & "(" & Round(x2, 2) & "," & Round(y2, 2) & "," & Round(Z2, 2) & ")" & _
        vbCrLf & "该视图不能生成正确的百格线,程序退出！"
End If
If scheck <> "" Then
    MsgBox scheck, vbOKOnly + vbInformation, "百格线工具"
    Exit Sub
End If

'Calculate the Grid limit, Programmed by Charles.Tang, Allright reserved

Dim x As Double
Dim y As Double
Dim dScale As Double
Dim dAngle As Double

x = oDrwView.xAxisData
y = oDrwView.yAxisData
dScale = oDrwView.Scale
dAngle = oDrwView.Angle


Dim Xmin, Xmax, Ymin, Ymax
Xmin = fViewSize(oDrwView, 0)
Xmax = fViewSize(oDrwView, 1)
Ymin = fViewSize(oDrwView, 2)
Ymax = fViewSize(oDrwView, 3)

'如果间距>0.5倍零件尺寸，提出警告信息
Dim Dis As Double
Dis = Val(txtDis.Text)

If Dis > ((Xmax - Xmin) / 2) / dScale Or Dis > ((Ymax - Ymin) / 2) / dScale Then
    If MsgBox("百格线间距 " & Dis & "，零件尺寸大约" & Round((Xmax - Xmin) / dScale, 0) & "x" & Round((Ymax - Ymin) / dScale, 0) & ",百格线间距过大！" & _
    vbCrLf & "确定[OK]继续，取消[Cancel]退出", vbOKCancel + vbInformation, "百格线工具") = vbCancel Then
    Exit Sub
    End If
End If


'左下和右上偏移
Xmin = Xmin - Val(txtLBX.Text)
Xmax = Xmax + Val(txtRUX.Text)
Ymin = Ymin - Val(txtLBY.Text)
Ymax = Ymax + Val(txtRUY.Text)


Dim AbsoluteCoordinatesMin(1)
Dim AbsoluteCoordinatesMax(1)
Dim RelativeCoordinatesMin(1)
Dim RelativeCoordinatesMax(1)


AbsoluteCoordinatesMin(0) = Xmin
AbsoluteCoordinatesMin(1) = Ymin

CatRelativeCoordinates oDrwView, AbsoluteCoordinatesMin, RelativeCoordinatesMin

AbsoluteCoordinatesMax(0) = Xmax
AbsoluteCoordinatesMax(1) = Ymax
CatRelativeCoordinates oDrwView, AbsoluteCoordinatesMax, RelativeCoordinatesMax

Dim MinX As Double
Dim MinY As Double
Dim MaxX As Double
Dim MaxY As Double


'坐标化整，间距的整数倍
    MinX = Dis * (RelativeCoordinatesMin(0) \ Dis) - Dis
    MinY = Dis * (RelativeCoordinatesMin(1) \ Dis) - Dis
    MaxX = Dis * (RelativeCoordinatesMax(0) \ Dis) + Dis
    MaxY = Dis * (RelativeCoordinatesMax(1) \ Dis) + Dis
    
'坐标线条数
Dim nx, ny As Integer
    nx = Abs((MaxX - MinX) \ Dis)   'X 轴条数
    ny = Abs((MaxY - MinY) \ Dis)   'Y 轴条数



Dim Di_H, Di_V
Dim Text_XYZ_H As String
Dim Text_XYZ_V As String

Di_H = 1
Di_V = 1
If Abs(Zn) = 1 Then

End If


If (x1 = 1) Then
    Text_XYZ_H = " X"
End If
If (x1 = -1) Then
    Text_XYZ_H = " X"
    Di_H = -1
End If

If (y1 = 1) Then
    Text_XYZ_H = " Y"
End If
If (y1 = -1) Then
    Text_XYZ_H = " Y"
    Di_H = -1
End If

If (Z1 = 1) Then
    Text_XYZ_H = " Z"
End If
If (Z1 = -1) Then
    Text_XYZ_H = " Z"
    Di_H = -1
End If


If (x2 = 1) Then
    Text_XYZ_V = " X"
End If
If (x2 = -1) Then
    Text_XYZ_V = " X"
    Di_V = -1
End If

If (y2 = 1) Then
    Text_XYZ_V = " Y"
End If
If (y2 = -1) Then
    Text_XYZ_V = " Y"
    Di_V = -1
End If

If (Z2 = 1) Then
    Text_XYZ_V = " Z"
End If
If (Z2 = -1) Then
    Text_XYZ_V = " Z"
    Di_V = -1
End If


Dim oDrwText As Object
Dim Line2D1 As Object
Dim i As Integer

Set oSel = oActDoc.Selection
oSel.Clear
'水平线

For i = 0 To ny

    Set Line2D1 = oDrwView.Factory2D.CreateLine(MinX - Dis / 3 - Val(txtFontSize.Value) / dScale, MinY + Dis * i, MaxX + Dis / 3, MinY + Dis * i)
    Set oDrwText = oDrwView.Texts.Add((MinY + Dis * i) * Di_V & Text_XYZ_V, MinX - Dis / 3 - Val(txtFontSize.Value) / dScale, MinY + Dis * i)
    oDrwText.AnchorPosition = catBaseCenter
    oDrwText.SetFontSize 0, 0, Val(txtFontSize.Text)
    Line2D1.Name = "Grid_H" & i
    oDrwText.Name = "Grid_Htxt" & i
    oSel.Add Line2D1
    oSel.Add oDrwText
    
Next

  '垂直线
For i = 0 To nx

    Set Line2D1 = oDrwView.Factory2D.CreateLine(MinX + Dis * i, MinY - Dis / 3, MinX + Dis * i, MaxY + Dis / 3)
    Set oDrwText = oDrwView.Texts.Add((MinX + Dis * i) * Di_H & Text_XYZ_H, MinX + Dis * i, MinY - Dis / 3)
    oDrwText.AnchorPosition = catCapCenter
    oDrwText.SetFontSize 0, 0, Val(txtFontSize.Text)
    Line2D1.Name = "Grid_V" & i
    oDrwText.Name = "Grid_Vtxt" & i
    oSel.Add Line2D1
    oSel.Add oDrwText
    
Next

'百格线颜色和线型
Select Case cbxLineColor.Text
    Case "黑色"
        oSel.VisProperties.SetRealColor 0, 0, 0, 1
    Case "蓝色"
        oSel.VisProperties.SetRealColor 0, 0, 255, 1
    Case "绿色"
        oSel.VisProperties.SetRealColor 0, 255, 0, 1
    Case "黄色"
        oSel.VisProperties.SetRealColor 255, 255, 0, 1
End Select

oSel.VisProperties.SetRealWidth 1, 1

Select Case cbxLineStyle.Text
    Case "细实线"
        oSel.VisProperties.SetRealLineType 1, 1
    Case "虚线"
        oSel.VisProperties.SetRealLineType 6, 1
End Select

oSel.Clear

'oActDoc.Sheets.ActiveSheet.Update
'
'Set oDrwText = oDrwView.Texts.Add("ProjectNormal is (" & Xn & "," & Yn & "," & Zn & ")", 0, -10 / dScale)
'Set oDrwText = oDrwView.Texts.Add("ProjectPlane is (" & x1 & "," & y1 & "," & Z1 & ")," & "(" & x2 & "," & y2 & "," & Z2 & ")", 0, -20 / dScale)
'Set oDrwText = oDrwView.Texts.Add("View (xAxis，yAxis) is (" & X & "," & Y & ")," & "View Scale is " & dScale & "," & "View Angle is " & dAngle, 0, -30 / dScale)
'Set oDrwText = oDrwView.Texts.Add("View LeftBottom corner is (" & Xmin & "," & Ymin & ")," & "View RightUpper corner is (" & Xmax & "," & Ymax & ")", 0, -40 / dScale)
'Set oDrwText = oDrwView.Texts.Add("View inside LeftBottom corner is (" & RelativeCoordinatesMin(0) & "," & RelativeCoordinatesMin(1) & ")," & "View Inside RightUpper corner is (" & RelativeCoordinatesMax(0) & "," & RelativeCoordinatesMax(1) & ")", 0, -50 / dScale)

'恢复命令执行前的设置
CATIA.RefreshDisplay = True
CATIA.DisplayFileAlerts = True

Set oDrwView = Nothing
Set oSel = Nothing
Set Line2D1 = Nothing
Set oDrwText = Nothing

End Sub


Private Sub UserForm_Initialize()

    cbxLineStyle.AddItem "细实线"
    cbxLineStyle.AddItem "虚线"
    
    cbxLineColor.AddItem "黑色"
    cbxLineColor.AddItem "蓝色"
    cbxLineColor.AddItem "绿色"
    cbxLineColor.AddItem "黄色"
ReadConf
End Sub
Function fViewSize(oView As Object, fla As Integer)  'fla must be 0,1,2,3
Dim oView1
Set oView1 = oView

Dim oXY(4)
oView1.Size oXY
'Dim Xmin, Xmax, Ymin, Ymax
'Xmin = oXY(0)
'Xmax = oXY(1)
'Ymin = oXY(2)
'Ymax = oXY(3)

fViewSize = oXY(fla)

'MsgBox "View LeftBottom corner is (" & Xmin & "," & Ymin & ")," & "View RightUpper corner is (" & Xmax & "," & Ymax & ")"
Set oView1 = Nothing
End Function
Private Sub CatRelativeCoordinates(CatDrawingView As Object, AbsoluteCoordinates(), RelativeCoordinates())

' Compute the the point coordinates in a view coordinates system according to the absolute coordinates
' Location, Angle and Scale factor of the view are take into account
RelativeCoordinates(0) = ((AbsoluteCoordinates(1) - CatDrawingView.yAxisData) * Sin(CatDrawingView.Angle) + (AbsoluteCoordinates(0) - CatDrawingView.xAxisData) * Cos(CatDrawingView.Angle)) / CatDrawingView.Scale2
RelativeCoordinates(1) = ((AbsoluteCoordinates(1) - CatDrawingView.yAxisData) * Cos(CatDrawingView.Angle) - (AbsoluteCoordinates(0) - CatDrawingView.xAxisData) * Sin(CatDrawingView.Angle)) / CatDrawingView.Scale2
RelativeCoordinates(0) = Round(RelativeCoordinates(0), 0)
RelativeCoordinates(1) = Round(RelativeCoordinates(1), 0)
End Sub
Private Sub txtDis_KeyPress(ByVal KeyANSI As MSForms.ReturnInteger)
 Select Case KeyANSI
          Case Asc("0") To Asc("9")
          'Case Asc("-")

          Case Else
            MsgBox "只能输入正整数！"
            
    End Select
End Sub
Private Sub txtFontSize_KeyPress(ByVal KeyANSI As MSForms.ReturnInteger)
 Select Case KeyANSI
          Case Asc("0") To Asc("9")
          Case Asc(".")

          Case Else
            MsgBox "只能输入数字！"
            
    End Select
End Sub
Private Sub txtLBX_KeyPress(ByVal KeyANSI As MSForms.ReturnInteger)
 Select Case KeyANSI
          Case Asc("0") To Asc("9")
          Case Asc("-")
          Case Asc(".")
          Case Else
            MsgBox "只能输入数字！"
            
    End Select
End Sub
Private Sub txtLBY_KeyPress(ByVal KeyANSI As MSForms.ReturnInteger)
 Select Case KeyANSI
          Case Asc("0") To Asc("9")
          Case Asc("-")
          Case Asc(".")
          Case Else
            MsgBox "只能输入数字！"
            
    End Select
End Sub
Private Sub txtRUX_KeyPress(ByVal KeyANSI As MSForms.ReturnInteger)
 Select Case KeyANSI
          Case Asc("0") To Asc("9")
          Case Asc("-")
          Case Asc(".")
          Case Else
            MsgBox "只能输入数字！"
            
    End Select
End Sub
Private Sub txtRUY_KeyPress(ByVal KeyANSI As MSForms.ReturnInteger)
 Select Case KeyANSI
          Case Asc("0") To Asc("9")
          Case Asc("-")
          Case Asc(".")
          Case Else
            MsgBox "只能输入数字！"
            
    End Select
End Sub
Private Sub ReadConf()
Dim sConfpath As String
sConfpath = oCATVBA_Folder.Path
Dim objFSO
Set objFSO = CreateObject("Scripting.FileSystemObject")
'MsgBox sConfpath
'MsgBox Dir(sConfpath)
If Not objFSO.FileExists(sConfpath & "\Conf1.ini") Then
Exit Sub
End If
'---读取配置文件---
'On Error Resume Next
    Dim read_OK ' As Long
    Dim read2 As String
    read2 = String(255, 0)
    read_OK = GetPrivateProfileString("百格线", "百格线间距", "100", read2, 256, sConfpath & "\Conf1.ini")
    txtDis.Text = read2
    read2 = String(255, 0)
    read_OK = GetPrivateProfileString("百格线", "线型", "细实线", read2, 256, sConfpath & "\Conf1.ini")
    cbxLineStyle.Text = read2
    read2 = String(255, 0)
    read_OK = GetPrivateProfileString("百格线", "颜色", "蓝色", read2, 256, sConfpath & "\Conf1.ini")
    cbxLineColor.Text = read2
    read2 = String(255, 0)
    read_OK = GetPrivateProfileString("百格线", "标注字号", "3.5", read2, 256, sConfpath & "\Conf1.ini")
    txtFontSize.Text = read2
    read2 = String(255, 0)
    read_OK = GetPrivateProfileString("百格线", "左下X-", "20", read2, 256, sConfpath & "\Conf1.ini")
    txtLBX.Text = read2
    read2 = String(255, 0)
    read_OK = GetPrivateProfileString("百格线", "左下Y-", "20", read2, 256, sConfpath & "\Conf1.ini")
    txtLBY.Text = read2
    read2 = String(255, 0)
    read_OK = GetPrivateProfileString("百格线", "右上X+", "20", read2, 256, sConfpath & "\Conf1.ini")
    txtRUX.Text = read2
    read2 = String(255, 0)
    read_OK = GetPrivateProfileString("百格线", "右上Y+", "20", read2, 256, sConfpath & "\Conf1.ini")
    txtRUY.Text = read2
End Sub
Private Sub SaveConf()
On Error Resume Next
Dim sConfpath As String
Dim write1 'As Long
    '参数一： Section Name (节的名称)。
    '参数二： 节下面的项目名称。
    '参数三： 项目的内容。
    '参数四： ini配置文件的路径名称。
              
sConfpath = oCATVBA_Folder.Path
write1 = WritePrivateProfileString("百格线", "百格线间距", txtDis.Text, sConfpath & "\Conf1.ini")
write1 = WritePrivateProfileString("百格线", "线型", cbxLineStyle.Text, sConfpath & "\Conf1.ini")
write1 = WritePrivateProfileString("百格线", "颜色", cbxLineColor.Text, sConfpath & "\Conf1.ini")
write1 = WritePrivateProfileString("百格线", "标注字号", txtFontSize.Text, sConfpath & "\Conf1.ini")
write1 = WritePrivateProfileString("百格线", "左下X-", txtLBX.Text, sConfpath & "\Conf1.ini")
write1 = WritePrivateProfileString("百格线", "左下Y-", txtLBY.Text, sConfpath & "\Conf1.ini")
write1 = WritePrivateProfileString("百格线", "右上X+", txtRUX.Text, sConfpath & "\Conf1.ini")
write1 = WritePrivateProfileString("百格线", "右上Y+", txtRUY.Text, sConfpath & "\Conf1.ini")
End Sub

Private Sub UserForm_Terminate()
SaveConf
End Sub
