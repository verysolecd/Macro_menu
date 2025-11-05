VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsControlHandler"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' 定义一个对主窗体的引用，用于回调
Private m_ParentForm As Object

' 为我们想处理的每一种控件类型都定义一个WithEvents变量
Private WithEvents Evt_Button As MSForms.CommandButton
Private WithEvents Evt_CheckBox As MSForms.CheckBox
Private WithEvents Evt_Label As MSForms.Label
Private WithEvents Evt_TextBox As MSForms.TextBox

' 通用的“附加”方法
' 它接收主窗体的引用和要处理的控件
Public Sub Attach(ByVal ParentForm As Object, ByVal control As MSForms.control)
    ' 保存对主窗体的引用，以便后续回调
    Set m_ParentForm = ParentForm
    
    ' 根据传入控件的类型，将其赋值给对应的WithEvents变量
    Select Case TypeName(control)
        Case "CommandButton"
            Set Evt_Button = control
        Case "CheckBox"
            Set Evt_CheckBox = control
        Case "Label"
            Set Evt_Label = control
        Case "TextBox"
            Set Evt_TextBox = control
    End Select
End Sub

' --- 事件处理程序 ---
' 无论哪个控件的事件被触发，都统一调用主窗体的公共方法

Private Sub Evt_Button_Click()
    ' 回调主窗体的OnControlEvent方法，传递事件类型和控件本身
    If Not m_ParentForm Is Nothing Then
        m_ParentForm.OnControlEvent "Click", Evt_Button
    End If
End Sub

Private Sub Evt_Label_Click()
    If Not m_ParentForm Is Nothing Then
        m_ParentForm.OnControlEvent "Click", Evt_Label
    End If
End Sub

Private Sub Evt_CheckBox_Change()
    If Not m_ParentForm Is Nothing Then
        m_ParentForm.OnControlEvent "Change", Evt_CheckBox
    End If
End Sub

Private Sub Evt_TextBox_Change()
    If Not m_ParentForm Is Nothing Then
        m_ParentForm.OnControlEvent "Change", Evt_TextBox
    End If
End Sub
