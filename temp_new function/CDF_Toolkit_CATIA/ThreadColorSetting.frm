VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ThreadColorSetting 
   Caption         =   "孔和螺纹颜色设置 | 臭豆腐工具箱CATIA版"
   ClientHeight    =   9570.001
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   10155
   OleObjectBlob   =   "ThreadColorSetting.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "ThreadColorSetting"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
#If VBA7 Then
Private Declare PtrSafe Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As LongPtr
Private Declare PtrSafe Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As LongPtr
#Else
Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Private Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long
#End If

Private Sub ReadConf()
Dim sConfpath As String
sConfpath = oCATVBA_Folder.Path
Dim objFSO
Set objFSO = CreateObject("Scripting.FileSystemObject")

If Not objFSO.FileExists(sConfpath & "\Conf1.ini") Then
Exit Sub
End If
'---读取配置文件---
On Error Resume Next
    Dim read_OK ' As Long
    Dim read2 As String
    Dim i As Integer
    Dim ctrDatxt As String
    Dim ctrDbtxt As String
    'Dim ctrColorName As String
    Dim ctrRtxt As String
    Dim ctrGtxt As String
    Dim ctrBtxt As String
    
    Dim ctrGeoRtxt As String
    Dim ctrGeoGtxt As String
    Dim ctrGeoBtxt As String
    Dim ctrGeoName As String
    
    For i = 1 To 11
        ctrDatxt = "txtD" & i & "a"
        ctrDbtxt = "txtD" & i & "b"
        'ctrColorName = "lblColor" & i
        ctrRtxt = "txtR" & i
        ctrGtxt = "txtG" & i
        ctrBtxt = "txtB" & i
        
        
        read2 = String(255, 0)
        read_OK = GetPrivateProfileString("孔和螺纹涂色", ctrDatxt, "False", read2, 256, sConfpath & "\Conf1.ini")
        Me.Controls.Item(ctrDatxt).Text = Val(read2)
        read2 = String(255, 0)
        read_OK = GetPrivateProfileString("孔和螺纹涂色", ctrDbtxt, "False", read2, 256, sConfpath & "\Conf1.ini")
        Me.Controls.Item(ctrDbtxt).Text = Val(read2)
        
        read2 = String(255, 0)
        read_OK = GetPrivateProfileString("孔和螺纹涂色", ctrRtxt, "False", read2, 256, sConfpath & "\Conf1.ini")
        Me.Controls.Item(ctrRtxt).Text = Val(read2)
        read2 = String(255, 0)
        read_OK = GetPrivateProfileString("孔和螺纹涂色", ctrGtxt, "False", read2, 256, sConfpath & "\Conf1.ini")
        Me.Controls.Item(ctrGtxt).Text = Val(read2)
        read2 = String(255, 0)
        read_OK = GetPrivateProfileString("孔和螺纹涂色", ctrBtxt, "False", read2, 256, sConfpath & "\Conf1.ini")
        Me.Controls.Item(ctrBtxt).Text = Val(read2)
    Next
    For i = 1 To 9
    ctrGeoRtxt = "txtGeoR" & i
    ctrGeoGtxt = "txtGeoG" & i
    ctrGeoBtxt = "txtGeoB" & i
    ctrGeoName = "txtGeoName" & i
    
    read2 = String(255, 0)
    read_OK = GetPrivateProfileString("几何体涂色", ctrGeoRtxt, "False", read2, 256, sConfpath & "\Conf1.ini")
    Me.Controls.Item(ctrGeoRtxt).Text = Val(read2)
    
    read2 = String(255, 0)
    read_OK = GetPrivateProfileString("几何体涂色", ctrGeoGtxt, "False", read2, 256, sConfpath & "\Conf1.ini")
    Me.Controls.Item(ctrGeoGtxt).Text = Val(read2)
    
    read2 = String(255, 0)
    read_OK = GetPrivateProfileString("几何体涂色", ctrGeoBtxt, "False", read2, 256, sConfpath & "\Conf1.ini")
    Me.Controls.Item(ctrGeoBtxt).Text = Val(read2)
    
    read2 = String(255, 0)
    read_OK = GetPrivateProfileString("几何体涂色", ctrGeoName, "", read2, 256, sConfpath & "\Conf1.ini")
    Me.Controls.Item(ctrGeoName).Text = read2
    Next
    
    read2 = String(255, 0)
    read_OK = GetPrivateProfileString("几何体涂色", "涂色色块显示名称", "False", read2, 256, sConfpath & "\Conf1.ini")
    Me.Controls.Item("chkShowName").Value = read2
    
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

Dim i As Integer
Dim ctrDatxt As String
Dim ctrDbtxt As String
Dim ctrRtxt As String
Dim ctrGtxt As String
Dim ctrBtxt As String

Dim ctrGeoRtxt As String
Dim ctrGeoGtxt As String
Dim ctrGeoBtxt As String
Dim ctrGeoName As String

For i = 1 To 11
ctrDatxt = "txtD" & i & "a"
ctrDbtxt = "txtD" & i & "b"
ctrRtxt = "txtR" & i
ctrGtxt = "txtG" & i
ctrBtxt = "txtB" & i
write1 = WritePrivateProfileString("孔和螺纹涂色", ctrDatxt, CStr(Me.Controls.Item(ctrDatxt).Text), sConfpath & "\Conf1.ini")
write1 = WritePrivateProfileString("孔和螺纹涂色", ctrDbtxt, CStr(Me.Controls.Item(ctrDbtxt).Text), sConfpath & "\Conf1.ini")
write1 = WritePrivateProfileString("孔和螺纹涂色", ctrRtxt, CStr(Me.Controls.Item(ctrRtxt).Text), sConfpath & "\Conf1.ini")
write1 = WritePrivateProfileString("孔和螺纹涂色", ctrGtxt, CStr(Me.Controls.Item(ctrGtxt).Text), sConfpath & "\Conf1.ini")
write1 = WritePrivateProfileString("孔和螺纹涂色", ctrBtxt, CStr(Me.Controls.Item(ctrBtxt).Text), sConfpath & "\Conf1.ini")
Next

For i = 1 To 9
ctrGeoRtxt = "txtGeoR" & i
ctrGeoGtxt = "txtGeoG" & i
ctrGeoBtxt = "txtGeoB" & i
ctrGeoName = "txtGeoName" & i
write1 = WritePrivateProfileString("几何体涂色", ctrGeoRtxt, CStr(Me.Controls.Item(ctrGeoRtxt).Text), sConfpath & "\Conf1.ini")
write1 = WritePrivateProfileString("几何体涂色", ctrGeoGtxt, CStr(Me.Controls.Item(ctrGeoGtxt).Text), sConfpath & "\Conf1.ini")
write1 = WritePrivateProfileString("几何体涂色", ctrGeoBtxt, CStr(Me.Controls.Item(ctrGeoBtxt).Text), sConfpath & "\Conf1.ini")
write1 = WritePrivateProfileString("几何体涂色", ctrGeoName, CStr(Me.Controls.Item(ctrGeoName).Text), sConfpath & "\Conf1.ini")
Next
write1 = WritePrivateProfileString("几何体涂色", "涂色色块显示名称", CStr(Me.Controls.Item("chkShowName").Value), sConfpath & "\Conf1.ini")
End Sub

Private Sub cmdCancel_Click()
Unload Me
End Sub

Private Sub cmdOK_Click()
SaveConf
MsgBox "重新运行命令才会生效！"
Unload Me
End Sub

Private Sub UserForm_Initialize()
        
        
On Error Resume Next
ReadConf
Dim i As Integer
Dim ctrColorName As String
Dim ctrRtxt As String
Dim ctrGtxt As String
Dim ctrBtxt As String

Dim ctrGeoColor As String
Dim ctrGeoRtxt As String
Dim ctrGeoGtxt As String
Dim ctrGeoBtxt As String

For i = 1 To 11
ctrColorName = "lblColor" & i
ctrRtxt = "txtR" & i
ctrGtxt = "txtG" & i
ctrBtxt = "txtB" & i
Me.Controls.Item(ctrColorName).BackColor = RGB(Val(Me.Controls.Item(ctrRtxt).Text), _
                                                Val(Me.Controls.Item(ctrGtxt).Text), _
                                                Val(Me.Controls.Item(ctrBtxt).Text))

Next
For i = 1 To 9
ctrGeoColor = "lblGeo" & i
ctrGeoRtxt = "txtGeoR" & i
ctrGeoGtxt = "txtGeoG" & i
ctrGeoBtxt = "txtGeoB" & i
ctrGeoName = "txtGeoName" & i
Me.Controls.Item(ctrGeoColor).BackColor = RGB(Val(Me.Controls.Item(ctrGeoRtxt).Text), _
                                                Val(Me.Controls.Item(ctrGeoGtxt).Text), _
                                                Val(Me.Controls.Item(ctrGeoBtxt).Text))
If chkShowName.Value = True Then
Me.Controls.Item(ctrGeoColor).Caption = Me.Controls.Item(ctrGeoName).Text
End If
Next

End Sub
Private Sub cmdPreview_Click()
Dim i As Integer
Dim ctrColorName As String
Dim ctrRtxt As String
Dim ctrGtxt As String
Dim ctrBtxt As String

Dim ctrGeoColor As String
Dim ctrGeoRtxt As String
Dim ctrGeoGtxt As String
Dim ctrGeoBtxt As String

For i = 1 To 11
ctrColorName = "lblColor" & i
ctrRtxt = "txtR" & i
ctrGtxt = "txtG" & i
ctrBtxt = "txtB" & i
Me.Controls.Item(ctrColorName).BackColor = RGB(Val(Me.Controls.Item(ctrRtxt).Text), _
                                                Val(Me.Controls.Item(ctrGtxt).Text), _
                                                Val(Me.Controls.Item(ctrBtxt).Text))
Next

For i = 1 To 9
ctrGeoColor = "lblGeo" & i
ctrGeoRtxt = "txtGeoR" & i
ctrGeoGtxt = "txtGeoG" & i
ctrGeoBtxt = "txtGeoB" & i
ctrGeoName = "txtGeoName" & i
Me.Controls.Item(ctrGeoColor).BackColor = RGB(Val(Me.Controls.Item(ctrGeoRtxt).Text), _
                                                Val(Me.Controls.Item(ctrGeoGtxt).Text), _
                                                Val(Me.Controls.Item(ctrGeoBtxt).Text))
If chkShowName.Value = True Then
Me.Controls.Item(ctrGeoColor).Caption = Me.Controls.Item(ctrGeoName).Text
Else
Me.Controls.Item(ctrGeoColor).Caption = ""
End If
Next
End Sub

Private Sub txtR1_KeyPress(ByVal KeyANSI As MSForms.ReturnInteger)
    Select Case KeyANSI
          Case Asc("0") To Asc("9")
          Case Else
            MsgBox "只能输入数字！数值为0-255之间的整数"
    End Select
End Sub
Private Sub txtR2_KeyPress(ByVal KeyANSI As MSForms.ReturnInteger)
    Select Case KeyANSI
          Case Asc("0") To Asc("9")
          Case Else
            MsgBox "只能输入数字！数值为0-255之间的整数"
    End Select
End Sub
Private Sub txtR3_KeyPress(ByVal KeyANSI As MSForms.ReturnInteger)
    Select Case KeyANSI
          Case Asc("0") To Asc("9")
          Case Else
            MsgBox "只能输入数字！数值为0-255之间的整数"
    End Select
End Sub
Private Sub txtR4_KeyPress(ByVal KeyANSI As MSForms.ReturnInteger)
    Select Case KeyANSI
          Case Asc("0") To Asc("9")
          Case Else
            MsgBox "只能输入数字！数值为0-255之间的整数"
    End Select
End Sub
Private Sub txtR5_KeyPress(ByVal KeyANSI As MSForms.ReturnInteger)
    Select Case KeyANSI
          Case Asc("0") To Asc("9")
          Case Else
            MsgBox "只能输入数字！数值为0-255之间的整数"
    End Select
End Sub
Private Sub txtR6_KeyPress(ByVal KeyANSI As MSForms.ReturnInteger)
    Select Case KeyANSI
          Case Asc("0") To Asc("9")
          Case Else
            MsgBox "只能输入数字！数值为0-255之间的整数"
    End Select
End Sub
Private Sub txtR7_KeyPress(ByVal KeyANSI As MSForms.ReturnInteger)
    Select Case KeyANSI
          Case Asc("0") To Asc("9")
          Case Else
            MsgBox "只能输入数字！数值为0-255之间的整数"
    End Select
End Sub
Private Sub txtR8_KeyPress(ByVal KeyANSI As MSForms.ReturnInteger)
    Select Case KeyANSI
          Case Asc("0") To Asc("9")
          Case Else
            MsgBox "只能输入数字！数值为0-255之间的整数"
    End Select
End Sub
Private Sub txtR9_KeyPress(ByVal KeyANSI As MSForms.ReturnInteger)
    Select Case KeyANSI
          Case Asc("0") To Asc("9")
          Case Else
            MsgBox "只能输入数字！数值为0-255之间的整数"
    End Select
End Sub
Private Sub txtR10_KeyPress(ByVal KeyANSI As MSForms.ReturnInteger)
    Select Case KeyANSI
          Case Asc("0") To Asc("9")
          Case Else
            MsgBox "只能输入数字！数值为0-255之间的整数"
    End Select
End Sub
Private Sub txtR11_KeyPress(ByVal KeyANSI As MSForms.ReturnInteger)
    Select Case KeyANSI
          Case Asc("0") To Asc("9")
          Case Else
            MsgBox "只能输入数字！数值为0-255之间的整数"
    End Select
End Sub
Private Sub txtG1_KeyPress(ByVal KeyANSI As MSForms.ReturnInteger)
    Select Case KeyANSI
          Case Asc("0") To Asc("9")
          Case Else
            MsgBox "只能输入数字！数值为0-255之间的整数"
    End Select
End Sub
Private Sub txtG2_KeyPress(ByVal KeyANSI As MSForms.ReturnInteger)
    Select Case KeyANSI
          Case Asc("0") To Asc("9")
          Case Else
            MsgBox "只能输入数字！数值为0-255之间的整数"
    End Select
End Sub
Private Sub txtG3_KeyPress(ByVal KeyANSI As MSForms.ReturnInteger)
    Select Case KeyANSI
          Case Asc("0") To Asc("9")
          Case Else
            MsgBox "只能输入数字！数值为0-255之间的整数"
    End Select
End Sub
Private Sub txtG4_KeyPress(ByVal KeyANSI As MSForms.ReturnInteger)
    Select Case KeyANSI
          Case Asc("0") To Asc("9")
          Case Else
            MsgBox "只能输入数字！数值为0-255之间的整数"
    End Select
End Sub
Private Sub txtG5_KeyPress(ByVal KeyANSI As MSForms.ReturnInteger)
    Select Case KeyANSI
          Case Asc("0") To Asc("9")
          Case Else
            MsgBox "只能输入数字！数值为0-255之间的整数"
    End Select
End Sub
Private Sub txtG6_KeyPress(ByVal KeyANSI As MSForms.ReturnInteger)
    Select Case KeyANSI
          Case Asc("0") To Asc("9")
          Case Else
            MsgBox "只能输入数字！数值为0-255之间的整数"
    End Select
End Sub
Private Sub txtG7_KeyPress(ByVal KeyANSI As MSForms.ReturnInteger)
    Select Case KeyANSI
          Case Asc("0") To Asc("9")
          Case Else
            MsgBox "只能输入数字！数值为0-255之间的整数"
    End Select
End Sub
Private Sub txtG8_KeyPress(ByVal KeyANSI As MSForms.ReturnInteger)
    Select Case KeyANSI
          Case Asc("0") To Asc("9")
          Case Else
            MsgBox "只能输入数字！数值为0-255之间的整数"
    End Select
End Sub
Private Sub txtG9_KeyPress(ByVal KeyANSI As MSForms.ReturnInteger)
    Select Case KeyANSI
          Case Asc("0") To Asc("9")
          Case Else
            MsgBox "只能输入数字！数值为0-255之间的整数"
    End Select
End Sub
Private Sub txtG10_KeyPress(ByVal KeyANSI As MSForms.ReturnInteger)
    Select Case KeyANSI
          Case Asc("0") To Asc("9")
          Case Else
            MsgBox "只能输入数字！数值为0-255之间的整数"
    End Select
End Sub
Private Sub txtG11_KeyPress(ByVal KeyANSI As MSForms.ReturnInteger)
    Select Case KeyANSI
          Case Asc("0") To Asc("9")
          Case Else
            MsgBox "只能输入数字！数值为0-255之间的整数"
    End Select
End Sub
Private Sub txtB1_KeyPress(ByVal KeyANSI As MSForms.ReturnInteger)
    Select Case KeyANSI
          Case Asc("0") To Asc("9")
          Case Else
            MsgBox "只能输入数字！数值为0-255之间的整数"
    End Select
End Sub
Private Sub txtB2_KeyPress(ByVal KeyANSI As MSForms.ReturnInteger)
    Select Case KeyANSI
          Case Asc("0") To Asc("9")
          Case Else
            MsgBox "只能输入数字！数值为0-255之间的整数"
    End Select
End Sub
Private Sub txtB3_KeyPress(ByVal KeyANSI As MSForms.ReturnInteger)
    Select Case KeyANSI
          Case Asc("0") To Asc("9")
          Case Else
            MsgBox "只能输入数字！数值为0-255之间的整数"
    End Select
End Sub
Private Sub txtB4_KeyPress(ByVal KeyANSI As MSForms.ReturnInteger)
    Select Case KeyANSI
          Case Asc("0") To Asc("9")
          Case Else
            MsgBox "只能输入数字！数值为0-255之间的整数"
    End Select
End Sub
Private Sub txtB5_KeyPress(ByVal KeyANSI As MSForms.ReturnInteger)
    Select Case KeyANSI
          Case Asc("0") To Asc("9")
          Case Else
            MsgBox "只能输入数字！数值为0-255之间的整数"
    End Select
End Sub
Private Sub txtB6_KeyPress(ByVal KeyANSI As MSForms.ReturnInteger)
    Select Case KeyANSI
          Case Asc("0") To Asc("9")
          Case Else
            MsgBox "只能输入数字！数值为0-255之间的整数"
    End Select
End Sub
Private Sub txtB7_KeyPress(ByVal KeyANSI As MSForms.ReturnInteger)
    Select Case KeyANSI
          Case Asc("0") To Asc("9")
          Case Else
            MsgBox "只能输入数字！数值为0-255之间的整数"
    End Select
End Sub
Private Sub txtB8_KeyPress(ByVal KeyANSI As MSForms.ReturnInteger)
    Select Case KeyANSI
          Case Asc("0") To Asc("9")
          Case Else
            MsgBox "只能输入数字！数值为0-255之间的整数"
    End Select
End Sub
Private Sub txtB9_KeyPress(ByVal KeyANSI As MSForms.ReturnInteger)
    Select Case KeyANSI
          Case Asc("0") To Asc("9")
          Case Else
            MsgBox "只能输入数字！数值为0-255之间的整数"
    End Select
End Sub
Private Sub txtB10_KeyPress(ByVal KeyANSI As MSForms.ReturnInteger)
    Select Case KeyANSI
          Case Asc("0") To Asc("9")
          Case Else
            MsgBox "只能输入数字！数值为0-255之间的整数"
    End Select
End Sub
Private Sub txtB11_KeyPress(ByVal KeyANSI As MSForms.ReturnInteger)
    Select Case KeyANSI
          Case Asc("0") To Asc("9")
          Case Else
            MsgBox "只能输入数字！数值为0-255之间的整数"
    End Select
End Sub
