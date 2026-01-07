Attribute VB_Name = "ThreadColor_3D"
#If VBA7 Then
Private Declare PtrSafe Function GetKeyState Lib "user32" (ByVal nVirtKey As Long) As Integer
#Else
Private Declare Function GetKeyState Lib "user32" (ByVal nVirtKey As Long) As Integer
#End If
#If VBA7 Then
Private Declare PtrSafe Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As LongPtr
Private Declare PtrSafe Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As LongPtr
#Else
Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Private Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long
#End If
Dim ResetCol As Boolean

'R1G1B1为孔的颜色, R2G2B2为螺纹面颜色
Dim R1 As Long
Dim G1 As Long
Dim B1 As Long
Dim R2 As Long
Dim G2 As Long
Dim B2 As Long
Dim confData(10, 4) As Double


Sub CATMain()
       
IntCATIA
On Error Resume Next
If TypeName(oActDoc) = "DrawingDocument" Then
        MsgBox "不能在工程制图模式下运行此命令!"
        Exit Sub
End If
ResetCol = False

If (GetKeyState(vbKeyShift) And &H8000&) Then
    ResetCol = True
End If
If (GetKeyState(vbKeyMenu) And &H8000&) Then
    ThreadColorSetting.Show vbModeless
    Exit Sub
End If

R1 = 0
G1 = 0
B1 = 255

R2 = 0
G2 = 100
B2 = 0

'给孔和螺纹颜色赋初值
confData(5, 2) = R1
confData(5, 3) = G1
confData(5, 4) = B1
confData(10, 2) = R2
confData(10, 3) = G2
confData(10, 4) = B2
    
ReadConf




CATIA.DisplayFileAlerts = False

Dim oProdxx
Set oProdxx = oActDoc.Product
oProdxx.ApplyWorkMode DESIGN_MODE '2
'-----------设置设计模式-------------

ThdColor oActDoc
CATIA.DisplayFileAlerts = True
End Sub
Sub ThdColor(oDocProd1)

If TypeName(oActDoc) = "ProductDocument" Then
Set oSel = oActDoc.Selection
ThdCol (oActDoc.Product)
ElseIf TypeName(oActDoc) = "PartDocument" Then
Set oSel = oActDoc.Selection
ThdCol2 oActDoc.Part
ElseIf TypeName(oActDoc) = "ProcessDocument" Then
Set oSel = oActDoc.Selection
ThdCol oActDoc.PPRDocument.Products.Item(1)
Else
MsgBox "此命令不能在该工作台中运行!" & vbCrLf & "请在装配设计或零件设计工作台运行此命令！"
Exit Sub
End If
End Sub

Sub ThdCol(oProduct)
oProduct.ApplyWorkMode DESIGN_MODE
oSel.Clear
oSel.Add oProduct
oSel.VisProperties.SetRealColor 255, 255, 255, 0
oSel.Clear

Dim j
For j = 1 To oProduct.Products.Count

If oProduct.Products.Item(j).Products.Count = 0 Then
On Error Resume Next
    oProduct.Products.Item(j).ApplyWorkMode DESIGN_MODE
    oSel.Add oProduct.Products.Item(j)
    oSel.VisProperties.SetRealColor 255, 255, 255, 0
    ThdCol2 (oProduct.Products.Item(j).ReferenceProduct.Parent.Part)
On Error GoTo 0
Else
ThdCol (oProduct.Products.Item(j))
End If

oSel.Clear
Next
oProduct.ApplyWorkMode DEFAULT_MODE
End Sub

Sub ThdCol2(oPart)


oSel.Clear
oSel.Add oPart
oSel.VisProperties.SetRealColor 255, 255, 255, 0

'*****************************
Dim oBodies, oBody, k, d
 
                    Set oBodies = oPart.Bodies
                        For k = 1 To oBodies.Count
                            Set oBody = oBodies.Item(k)
                            Dim shapes1 'As Shapes
                            Dim m As Integer
                            Set shapes1 = oBody.Shapes
                            For m = 1 To shapes1.Count
                                    If TypeName(shapes1.Item(m)) = "Hole" Then
                                             oSel.Add shapes1.Item(m)
                                            If shapes1.Item(m).ThreadingMode = 0 Then                        '孔,且带螺纹
                                                        If ResetCol = False Then
                                                            d = Val(shapes1.Item(m).ThreadDiameter.ValueAsString)
                                                            defthreadcolor d
                                                            oSel.VisProperties.SetRealColor R2, G2, B2, 1
                                                            
                                                        Else
                                                            oSel.VisProperties.SetRealColor 210, 210, 255, 1
                                                            oSel.VisProperties.SetRealOpacity 255, 1
                                                        End If
                                             Else                                                               '孔,但不带螺纹
                                                        If ResetCol = False Then
                                                            d = Val(shapes1.Item(m).Diameter.ValueAsString)
                                                            defholecolor d
                                                            oSel.VisProperties.SetRealColor R1, G1, B1, 1
                                                        Else
                                                            oSel.VisProperties.SetRealColor 210, 210, 255, 1
                                                            oSel.VisProperties.SetRealOpacity 255, 1
                                                        End If
                                             End If
                                            oSel.Clear
                                     End If
                                    If TypeName(shapes1.Item(m)) = "Thread" Then                                '螺纹面
                                              oSel.Add shapes1.Item(m)
                                                       If ResetCol = False Then
                                                            d = Val(shapes1.Item(m).Diameter) '.ValueAsString)
                                                            defthreadcolor d
                                                            oSel.VisProperties.SetRealColor R2, G2, B2, 1
                                                       Else
                                                           oSel.VisProperties.SetRealColor 210, 210, 255, 1
                                                           oSel.VisProperties.SetRealOpacity 255, 1
                                                       End If
                                     End If
                                     oSel.Clear
                            Next
                        Next


End Sub

Private Sub ReadConf()
Dim sConfpath As String
sConfpath = oCATVBA_Folder.Path
Dim objFSO
Set objFSO = CreateObject("Scripting.FileSystemObject")

If Not objFSO.FileExists(sConfpath & "\Conf1.ini") Then
Exit Sub
End If
'---读取配置文件---
'On Error Resume Next
    Dim read_OK ' As Long
    Dim read2 As String
    Dim i As Integer
    Dim ctrDatxt As String
    Dim ctrDbtxt As String
    'Dim ctrColorName As String
    Dim ctrRtxt As String
    Dim ctrGtxt As String
    Dim ctrBtxt As String
'    '给孔和螺纹颜色赋初值
'    confData(5, 2) = R1
'    confData(5, 3) = G1
'    confData(5, 4) = B1
'    confData(10, 2) = R2
'    confData(10, 3) = G2
'    confData(10, 4) = B2
    
    For i = 1 To 11
        ctrDatxt = "txtD" & i & "a"
        ctrDbtxt = "txtD" & i & "b"
        'ctrColorName = "lblColor" & i
        ctrRtxt = "txtR" & i
        ctrGtxt = "txtG" & i
        ctrBtxt = "txtB" & i
        
        
        read2 = String(255, 0)
        read_OK = GetPrivateProfileString("孔和螺纹涂色", ctrDatxt, "False", read2, 256, sConfpath & "\Conf1.ini")
        confData(i - 1, 0) = Val(read2)
        read2 = String(255, 0)
        read_OK = GetPrivateProfileString("孔和螺纹涂色", ctrDbtxt, "False", read2, 256, sConfpath & "\Conf1.ini")
        confData(i - 1, 1) = Val(read2)
        
        read2 = String(255, 0)
        read_OK = GetPrivateProfileString("孔和螺纹涂色", ctrRtxt, "False", read2, 256, sConfpath & "\Conf1.ini")
        confData(i - 1, 2) = Val(read2)
        read2 = String(255, 0)
        read_OK = GetPrivateProfileString("孔和螺纹涂色", ctrGtxt, "False", read2, 256, sConfpath & "\Conf1.ini")
        confData(i - 1, 3) = Val(read2)
        read2 = String(255, 0)
        read_OK = GetPrivateProfileString("孔和螺纹涂色", ctrBtxt, "False", read2, 256, sConfpath & "\Conf1.ini")
        confData(i - 1, 4) = Val(read2)
    Next

End Sub
Private Sub defthreadcolor(ByVal d As Double)
On Error Resume Next
R2 = 0
G2 = 100
B2 = 0
R2 = confData(10, 2)
G2 = confData(10, 3)
B2 = confData(10, 4)
Dim i As Integer
For i = 6 To 9
    If (confData(i, 0) <= d) And (d <= confData(i, 1)) Then
    R2 = confData(i, 2)
    G2 = confData(i, 3)
    B2 = confData(i, 4)
    Exit For
    End If
Next

End Sub
Private Sub defholecolor(ByVal d As Double)
On Error Resume Next
R1 = 0
G1 = 0
B1 = 255
R1 = confData(5, 2)
G1 = confData(5, 3)
B1 = confData(5, 4)
Dim i As Integer
For i = 0 To 4
    If (confData(i, 0) <= d) And (d <= confData(i, 1)) Then
    R1 = confData(i, 2)
    G1 = confData(i, 3)
    B1 = confData(i, 4)
    Exit For
    End If
Next
End Sub
