VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ProductObserver"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'VERSION 1.0 CLASS
'BEGIN
'  MultiUse = -1  'True
'End
'Attribute VB_Name = "ProductObserver"
'Attribute VB_GlobalNameSpace = False
'Attribute VB_Creatable = False
'Attribute VB_PredeclaredId = False
'Attribute VB_Exposed = False
Option Explicit

' 事件声明
Public Event ProductChanged()
' 私有变量存储当前产品
Private mProduct As Object

' 属性：当前产品
Public Property Get CurrentProduct() As Object
    Set CurrentProduct = mProduct
End Property
Public Property Set CurrentProduct(ByVal newProduct As Object)
    ' 设置新产品并触发事件
    Set mProduct = newProduct
    RaiseEvent ProductChanged
End Property
' 初始化
Private Sub Class_Initialize()
    Set mProduct = Nothing
End Sub
' 清理
Private Sub Class_Terminate()
    Set mProduct = Nothing
End Sub
