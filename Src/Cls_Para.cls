VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Cls_Para"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'我是一个对参数对象赋予属性的类
Private pName As String  '参数名称
Private pDesc As String  '参数描述
Private pType As String '参数类型
Private pStr As Variant '参数字符段
Private pTarget As Object '参数目标对象
Private pobj As Object   '参数本身

Private Sub Class_Initialize()
    Me.Desc = ""
End Sub
Public Sub Reset()
    Me.Name = ""
    Me.iType = ""
    Me.str = ""
    Me.Desc = ""
    Set Me.obj = Nothing
    Set Me.Target = Nothing
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
Public Property Get Desc() As Variant
    Desc = pDesc
End Property
Public Property Let Desc(ByVal value As Variant)
    pDesc = value
End Property
Public Property Get str() As Variant
    str = pStr
End Property
Public Property Let str(ByVal value As Variant)
    pStr = value
End Property
Public Property Get Target()
   Set Target = pTarget
End Property
Public Property Set Target(value)
   Set pTarget = value
End Property
Public Property Get obj()
   Set obj = pobj
End Property
Public Property Set obj(value)
   Set pobj = value
End Property
Public Sub SetNT(ByVal newName As String, ByVal newType As String)
    Me.Name = newName
    Me.iType = newType
End Sub
