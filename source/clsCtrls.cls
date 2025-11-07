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

Public WithEvents Txt As MSForms.textbox
Attribute Txt.VB_VarHelpID = -1
Public ohint As String

' 类模块：clsControlConfig（简化版）
Public ControlType As String ' 控件类型：Label/CommandButton/TextBox/OptionButton/CheckBox
Public Name As String ' 唯一名称
Public Caption As String ' 显示文本
Public DefaultValue As Variant ' 默认值（输入/选择类）
Public MultiLine As Boolean ' 文本框是否多行（仅TextBox用）
Private mform

Private Sub txt_Gotfocus()

If Txt.Text = ohint Then Txt.Text = "": Txt.ForeColor = &H0

End Sub
Private Sub txt_Lostfocus()

If Txt.Text = "" Then Txt.Text = ohint: Txt.ForeColor = &H808080

End Sub
