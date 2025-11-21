VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsCtrls"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Public WithEvents txt As MSForms.TextBox
Attribute txt.VB_VarHelpID = -1
Public ohint As String


Public Caption As String ' 显示文本
Public DefaultValue As Variant ' 默认值（输入/选择类）
Public MultiLine As Boolean ' 文本框是否多行（仅TextBox用）
Private mform

'我是一个对参数对象赋予属性的类
Private pName As String  '参数名称
Private pDesc As String  '参数描述
Private pType As String '参数类型
Private pStr As Variant '参数字符段
Private pTarget As Object '参数目标对象
Private pobj As Object   '参数本身


Private Sub Class_Initialize(ByVal ary As Variant)
  
End Sub


Public Sub Reset()
    Me.Name = ""
    Me.iType = ""

End Sub
Public Property Get Name() As String
    Name = pName
End Property
Public Property Let Name(ByVal value As String)
    pName = value
End Property
Public Property Get iType() As String
    iType = pType
End Property
Public Property Let iType(ByVal value As String)
    pType = value
End Property


Public Sub SetNT(ByVal newName As String, ByVal newType As String)
    Me.Name = newName
    Me.iType = newType
End Sub


Private Sub txt_Gotfocus()
If txt.Text = ohint Then txt.Text = "": txt.ForeColor = &H0
End Sub
Private Sub txt_Lostfocus()
If txt.Text = "" Then txt.Text = ohint: txt.ForeColor = &H808080
End Sub

Private init
