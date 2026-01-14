VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cls_SaveInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

    Public level As Integer      ' 产品在装配体中的层级（根产品为1级）
    Public prod As Product       ' 产品对象（Product/Part）
    Public partNumber As String  ' 零件号（冗余存储，方便后续扩展）
Option Explicit '强制声明变量，避免语法错误

' ========== 自定义类：保存产品层级和对象信息 ==========

Private Sub Class_Initialize()
   
End Sub



