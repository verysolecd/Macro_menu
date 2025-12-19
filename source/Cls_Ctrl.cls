VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Cls_Ctrl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False


Option Explicit

' 声明带事件的按钮对象
Public WithEvents ControlBtn As MSForms.CommandButton
Attribute ControlBtn.VB_VarHelpID = -1
Public ParentFrm As Object ' 指向当前的 UserForm
Public ControlID As String ' 用于识别是哪个按钮（如 "btnOK" 或 "btnCancel"）

Private Sub ControlBtn_Click()
    ' 1. 将点击的按钮 ID 存入窗体的 Tag 属性，作为返回值标识
    ParentFrm.Tag = ControlID
    ' 2. 隐藏窗体（注意：不是 Unload），这样 mdl2wd 才能继续读取其他控件的值
    ParentFrm.Hide
End Sub


