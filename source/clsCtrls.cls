VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsCtrls"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

' 声明一个 WithEvents 变量，用于监听控件的事件
' 使用 MSForms.Control 可以监听不同类型控件的共同事件，如 Click


Public WithEvents txt As MSForms.TextBox
Attribute txt.VB_VarHelpID = -1
Private WithEvents m_ctrl As MSForms.Control
Attribute m_ctrl.VB_VarHelpID = -1
Public ohint As String

Public Caption As String ' 显示文本
Public DefaultValue As Variant ' 默认值（输入/选择类）
Public MultiLine As Boolean ' 文本框是否多行（仅TextBox用）
Private mform

Option Explicit

' 这个属性用于将动态创建的控件传递给类实例
Public Property Set Control(objCtrl As MSForms.Control)
    Set m_ctrl = objCtrl
End Property
Public Property Get Control(objCtrl As MSForms.Control)
    Set m_ctrl = objCtrl
End Property

' 这是控件的 Click 事件处理过程
' 当被监听的控件被点击时，这段代码会自动执行
Private Sub m_ctrl_Click()
    Dim txt_TM As MSForms.TextBox
    
    ' 尝试在窗体上找到名为 "txt_TM" 的文本框
    On Error Resume Next
    Set txt_TM = m_ctrl.Parent.controls("txt_TM")
    On Error GoTo 0
    
    ' 根据控件的类型和名称执行不同的操作
    Select Case TypeName(m_ctrl)
        Case "CommandButton"
            MsgBox "动态按钮 '" & m_ctrl.Name & "' 被点击了！", vbInformation
            
           Call mBtn_click(m_ctrl.Name)
        
        Case "CheckBox"
        
            If Not txt_TM Is Nothing Then
                ' 在这里实现你原来的逻辑：用 CheckBox 控制 TextBox 的 Enabled 属性
                txt_TM.Enabled = m_ctrl.value
                Debug.Print "复选框 '" & m_ctrl.Name & "' 的状态改变为 " & m_ctrl.value & _
                            "，文本框 'txt_TM' 的 Enabled 属性已同步。"
            Else
                MsgBox "复选框 '" & m_ctrl.Name & "' 被点击，但未找到名为 'txt_TM' 的文本框。", vbExclamation
            End If
    End Select
    
    
End Sub

Function mBtn_click(str1)


End Function

Function mChk_Click()


End Function

Private Sub txt_Gotfocus()
If txt.Text = ohint Then txt.Text = "": txt.ForeColor = &H0
End Sub
Private Sub txt_Lostfocus()
If txt.Text = "" Then txt.Text = ohint: txt.ForeColor = &H808080
End Sub
