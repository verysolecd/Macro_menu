VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Class_para"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'我是一个对参数对象赋予属性的类
Private pName As String
Private pDesc As String
Private pType As String
Private pStr As Variant
Private pTarget As Object
Private pobj As Object

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
Public Property Let Name(ByVal Value As String)
    pName = Value
End Property
Public Property Get iType() As String
    iType = pType
End Property
Public Property Let iType(ByVal Value As String)
    pType = Value
End Property
Public Property Get Desc() As Variant
    Desc = pDesc
End Property
Public Property Let Desc(ByVal Value As Variant)
    pDesc = Value
End Property
Public Property Get str() As Variant
    str = pStr
End Property
Public Property Let str(ByVal Value As Variant)
    pStr = Value
End Property
Public Property Get Target()
   Set Target = pTarget
End Property
Public Property Set Target(Value)
   Set pTarget = Value
End Property

Public Property Get obj()
   Set obj = pobj
End Property
Public Property Set obj(Value)
   Set pobj = Value
End Property
Public Sub SetNT(ByVal newName As String, ByVal newType As String)
    Me.Name = newName
    Me.iType = newType
End Sub
