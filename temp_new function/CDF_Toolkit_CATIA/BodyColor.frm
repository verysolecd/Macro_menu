VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} BodyColor 
   Caption         =   "几何体涂色"
   ClientHeight    =   7830
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   2310
   OleObjectBlob   =   "BodyColor.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "BodyColor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

#If VBA7 Then
Private Declare PtrSafe Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As LongPtr
Private Declare PtrSafe Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As LongPtr
Private Declare PtrSafe Function GetKeyState Lib "user32" (ByVal nVirtKey As Long) As Integer
#Else
Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Private Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long
Private Declare Function GetKeyState Lib "user32" (ByVal nVirtKey As Long) As Integer
#End If
Dim boolShowGeoName As Boolean '是否在几何体色块上显示名称
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
    Dim ctrbtmColor As String
    Dim Ri
    Dim Gi
    Dim Bi
    Dim ctrGeoRtxt As String
    Dim ctrGeoGtxt As String
    Dim ctrGeoBtxt As String
    Dim ctrGeoName As String
    
    read2 = String(255, 0)
    read_OK = GetPrivateProfileString("几何体涂色", "涂色色块显示名称", "False", read2, 256, sConfpath & "\Conf1.ini")
    boolShowGeoName = CBool(read2)
   
    For i = 1 To 9
    ctrbtmColor = "btmColor" & i
    ctrGeoRtxt = "txtGeoR" & i
    ctrGeoGtxt = "txtGeoG" & i
    ctrGeoBtxt = "txtGeoB" & i
    ctrGeoName = "txtGeoName" & i
    
    read2 = String(255, 0)
    read_OK = GetPrivateProfileString("几何体涂色", ctrGeoRtxt, "False", read2, 256, sConfpath & "\Conf1.ini")
    Ri = Val(read2)
    
    read2 = String(255, 0)
    read_OK = GetPrivateProfileString("几何体涂色", ctrGeoGtxt, "False", read2, 256, sConfpath & "\Conf1.ini")
    Gi = Val(read2)
    
    read2 = String(255, 0)
    read_OK = GetPrivateProfileString("几何体涂色", ctrGeoBtxt, "False", read2, 256, sConfpath & "\Conf1.ini")
    Bi = Val(read2)
    Me.Controls.Item(ctrbtmColor).BackColor = RGB(Ri, Gi, Bi)
    If boolShowGeoName = True Then
        read2 = String(255, 0)
        read_OK = GetPrivateProfileString("几何体涂色", ctrGeoName, "", read2, 256, sConfpath & "\Conf1.ini")
        Me.Controls.Item(ctrbtmColor).Caption = read2
        Me.Controls.Item(ctrbtmColor).ControlTipText = Me.Controls.Item(ctrbtmColor).Caption
    End If
    Next

    
End Sub



Private Sub btmColor1_Click()
On Error Resume Next
CATIA.StartCommand "Body_Color_3D"

If (GetKeyState(vbKeyShift) And &H8000&) Then
    setCol "Face", btmColor1.BackColor
Else
setCol "Body", btmColor1.BackColor
End If

End Sub
Private Sub btmColor2_Click()
On Error Resume Next
CATIA.StartCommand "Body_Color_3D"

If (GetKeyState(vbKeyShift) And &H8000&) Then
    setCol "Face", btmColor2.BackColor
Else
setCol "Body", btmColor2.BackColor
End If

End Sub
Private Sub btmColor3_Click()
On Error Resume Next
CATIA.StartCommand "Body_Color_3D"

If (GetKeyState(vbKeyShift) And &H8000&) Then
    setCol "Face", btmColor3.BackColor
Else
setCol "Body", btmColor3.BackColor
End If

End Sub
Private Sub btmColor4_Click()
On Error Resume Next
CATIA.StartCommand "Body_Color_3D"

If (GetKeyState(vbKeyShift) And &H8000&) Then
    setCol "Face", btmColor4.BackColor
Else
setCol "Body", btmColor4.BackColor
End If

End Sub
Private Sub btmColor5_Click()
On Error Resume Next
CATIA.StartCommand "Body_Color_3D"

If (GetKeyState(vbKeyShift) And &H8000&) Then
    setCol "Face", btmColor5.BackColor
Else
setCol "Body", btmColor5.BackColor
End If

End Sub
Private Sub btmColor6_Click()
On Error Resume Next
CATIA.StartCommand "Body_Color_3D"

If (GetKeyState(vbKeyShift) And &H8000&) Then
    setCol "Face", btmColor6.BackColor
Else
setCol "Body", btmColor6.BackColor
End If

End Sub
Private Sub btmColor7_Click()
On Error Resume Next
CATIA.StartCommand "Body_Color_3D"

If (GetKeyState(vbKeyShift) And &H8000&) Then
    setCol "Face", btmColor7.BackColor
Else
setCol "Body", btmColor7.BackColor
End If

End Sub
Private Sub btmColor8_Click()
On Error Resume Next
CATIA.StartCommand "Body_Color_3D"

If (GetKeyState(vbKeyShift) And &H8000&) Then
    setCol "Face", btmColor8.BackColor
Else
setCol "Body", btmColor8.BackColor
End If

End Sub
Private Sub btmColor9_Click()
On Error Resume Next
CATIA.StartCommand "Body_Color_3D"

If (GetKeyState(vbKeyShift) And &H8000&) Then
    setCol "Face", btmColor9.BackColor
Else
setCol "Body", btmColor9.BackColor
End If

End Sub


Private Sub cmdClearColor_Click()
On Error Resume Next
CATIA.StartCommand "Body_Color_3D"
Dim f As Boolean
If (GetKeyState(vbKeyShift) And &H8000&) Then
    f = True
End If
ClearCol oActDoc.Product, f

End Sub

Private Sub cmdColorsetting_Click()
On Error Resume Next
CATIA.StartCommand "Body_Color_3D"
ThreadColorSetting.Show vbModeless
End Sub

Private Sub cmdHideBody_Click()
On Error Resume Next
CATIA.StartCommand "Body_Color_3D"
HideBodies
End Sub

Private Sub cmdHideDatum_Click()
On Error Resume Next
CATIA.StartCommand "Body_Color_3D"
Hide_Datums_3D.CATMain
End Sub

Private Sub cmdRandomColor_Click()
On Error Resume Next
CATIA.StartCommand "Body_Color_3D"
Random_Color_3D.CATMain
End Sub

Private Sub cmdReverse_Click()

On Error Resume Next
CATIA.StartCommand "Body_Color_3D"
Body_ShowReverse_3D.CATMain
alreadysh.RemoveAll
End Sub

Private Sub cmdShowBody_Click()

On Error Resume Next
CATIA.StartCommand "Body_Color_3D"
IntCATIA
Body_ShowAll_3D.CATMain
End Sub

Private Sub cmdThreadHoleColor_Click()
On Error Resume Next
CATIA.StartCommand "Body_Color_3D"
ThreadColor_3D.CATMain
End Sub

Private Sub UserForm_Initialize()
IntCATIA
If TypeName(oActDoc) <> "ProductDocument" Then
   If TypeName(oActDoc) <> "PartDocument" Then
        MsgBox "当前文件不是CATIA零件或产品 !" & vbCrLf & "此命令只能在零件或产品文件中运行!"
        Exit Sub
   End If
End If
ReadConf
On Error Resume Next

Select Case TypeName(oActDoc)
    Case "PartDocument"
    Case "ProductDocument"
    Case Else
        MsgBox "不能在这个工作环境下运行此命令!", vbInformation, "臭豆腐工具箱CATIA版"
        Unload Me
End Select


End Sub
Sub delay(T As Single)
Dim time1 As Single
time1 = Timer
Do
DoEvents
Loop While Timer - time1 < T

End Sub
Private Function getR(ByVal i As Long)
getR = (i Mod 65536) Mod 256
End Function
Private Function getG(ByVal i As Long)
getG = (i Mod 65536) \ 256
End Function
Private Function getB(ByVal i As Long)
getB = i \ 65536
End Function

Private Sub solidCol(oBody, R, G, b)
Dim shapes1 'As Shapes
Dim m As Integer
Dim oSel
Set oSel = oActDoc.Selection
Set shapes1 = oBody.Shapes
For m = 1 To shapes1.Count
 If TypeName(shapes1.Item(m)) = "Solid" Then
    oSel.Add shapes1.Item(m)
        oSel.VisProperties.SetRealColor R, G, b, 1
    oSel.Clear
 End If
Next
End Sub
Private Sub setCol(objType As String, i As Long, Optional eR As Boolean = False)
On Error Resume Next
Dim objSel 'As Selection
Set objSel = oActDoc.Selection
objSel.Clear
Dim InputObjectType(0)
InputObjectType(0) = objType
Dim Status
Dim FlagFinish
FlagFinish = 0
Do Until FlagFinish = 1

    Status = objSel.SelectElement3(InputObjectType, "Select the " & InputObjectType(0) & ". Press Esc to finish the selection.", 0, 1, 1)
    If (Status = "Cancel") Then
        FlagFinish = 1
        objSel.Clear
        Exit Sub
    End If
    objSel.VisProperties.SetRealColor getR(i), getG(i), getB(i), 1
    If objType = "Body" And objSel.Count2 <> 0 Then
        Dim j As Integer
        For j = 1 To objSel.Count2
        solidCol objSel.Item(j).Value, getR(i), getG(i), getB(i)
        Next
    End If
    objSel.Clear
Loop
End Sub
Sub ClearCol(oProduct, Optional ByVal face1 As Boolean = False)
oProduct.ApplyWorkMode DESIGN_MODE
Set oSel = oActDoc.Selection
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
    ClearCol2 (oProduct.Products.Item(j).ReferenceProduct.Parent.Part), face1
On Error GoTo 0
Else
ClearCol oProduct.Products.Item(j), face1
End If
oSel.Clear
Next
oProduct.ApplyWorkMode DEFAULT_MODE
End Sub
Sub ClearCol2(oPart, Optional ByVal face1 As Boolean = False)
Dim R, G, b
Set oSel = oActDoc.Selection
oSel.Clear
oSel.Add oPart
oSel.VisProperties.SetRealColor 255, 255, 255, 0
If face1 = True Then
'*****************************
Dim oBodies, oBody, k
 
                    Set oBodies = oPart.Bodies
                        For k = 1 To oBodies.Count
                            Set oBody = oBodies.Item(k)
                            oSel.Add oBody

                                oSel.VisProperties.SetRealColor 210, 210, 255, 1
                                oSel.VisProperties.SetRealOpacity 255, 1

                            oSel.Clear

                            Dim shapes1 'As Shapes
                            Dim m As Integer
                            Set shapes1 = oBody.Shapes
                            For m = 1 To shapes1.Count
                             If TypeName(shapes1.Item(m)) = "Solid" Then
                                oSel.Add shapes1.Item(m)

                                    oSel.VisProperties.SetRealColor 210, 210, 255, 1
                                    oSel.VisProperties.SetRealOpacity 255, 1

                                oSel.Clear
                            End If
                            Next
                        Next
End If
End Sub
Public Sub HideBodies(Optional h1 As Boolean = True)

'On Error Resume Next
'running = 1
Dim objSel 'As Selection
Set objSel = oActDoc.Selection
objSel.Clear
Dim InputObjectType(0)
InputObjectType(0) = "Body"
Dim Status
Dim FlagFinish
FlagFinish = 0
Do Until FlagFinish = 1
    'objSel.Clear
    Status = objSel.SelectElement3(InputObjectType, "Select the " & InputObjectType(0) & ". Press Esc to finish the selection.", 0, 1, 1)
    If (Status = "Cancel") Then
        FlagFinish = 1
        objSel.Clear
        Exit Sub
    End If
    If h1 = True Then
    objSel.VisProperties.SetShow 1 '0show,1 noshow
    Else
    objSel.VisProperties.SetShow 0 '0show,1 noshow
    End If

Loop
End Sub

