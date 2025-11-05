VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmCustomDialog 
   Caption         =   "frmCustomDialog"
   ClientHeight    =   2500
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   3500
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmCustomDialog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' 用一个集合来存储所有的事件处理器对象，防止它们被释放
Private m_ControlHandlers As Collection

' 窗体初始化
Private Sub UserForm_Initialize()
    Set m_ControlHandlers = New Collection
End Sub

' 窗体终止，清理资源
Private Sub UserForm_Terminate()
    Set m_ControlHandlers = Nothing
End Sub

' --- 公共的中央事件调度方法 ---
Public Sub OnControlEvent(ByVal eventType As String, ByVal ctrl As MSForms.control)
    Debug.Print "Event: '" & eventType & "' from Control: '" & ctrl.Name & "'"
    
    Select Case eventType
        Case "Click"
            If ctrl.Name = "btnOK" Then
                MsgBox "你点击了OK按钮！", vbInformation, "事件响应"
                Unload Me
            ElseIf ctrl.Name = "btnCancel" Then
                MsgBox "你点击了Cancel按钮！", vbInformation, "事件响应"
                Unload Me
            Else
                MsgBox "你点击了标签: '" & ctrl.Name & "'", vbInformation, "事件响应"
            End If
            
        Case "Change"
            Select Case TypeName(ctrl)
                Case "CheckBox"
                    MsgBox "复选框 '" & ctrl.Name & "' 的状态变为: " & ctrl.Value, vbInformation, "事件响应"
                Case "TextBox"
                    Me.caption = "文本已变更为: " & ctrl.Text
            End Select
    End Select
End Sub


' --- 用于动态添加控件的公共方法 ---

Public Sub AddButton(name As String, caption As String, left As Single, top As Single, width As Single, height As Single)
    Dim btn As MSForms.CommandButton
    Set btn = Me.Controls.Add("Forms.CommandButton.1", name, True)
    With btn
        .caption = caption
        .left = left
        .top = top
        .width = width
        .height = height
    End With
    CreateHandler btn
End Sub

Public Sub AddCheckBox(name As String, caption As String, left As Single, top As Single)
    Dim chk As MSForms.CheckBox
    Set chk = Me.Controls.Add("Forms.CheckBox.1", name, True)
    With chk
        .caption = caption
        .left = left
        .top = top
        .width = Len(caption) * 6 + 20 ' 自动宽度
        .height = 20
    End With
    CreateHandler chk
End Sub

Public Sub AddTextBox(name As String, text As String, left As Single, top As Single, width As Single, height As Single)
    Dim txt As MSForms.TextBox
    Set txt = Me.Controls.Add("Forms.TextBox.1", name, True)
    With txt
        .text = text
        .left = left
        .top = top
        .width = width
        .height = height
    End With
    CreateHandler txt
End Sub

Public Sub AddLabel(name As String, caption As String, left As Single, top As Single)
    Dim lbl As MSForms.Label
    Set lbl = Me.Controls.Add("Forms.Label.1", name, True)
    With lbl
        .caption = caption
        .left = left
        .top = top
        .width = Len(caption) * 6 + 10 ' 自动宽度
        .height = 20
    End With
    CreateHandler lbl
End Sub

' --- 核心私有方法 ---

Private Sub CreateHandler(control As MSForms.control)
    Dim handler As New clsControlHandler
    handler.Attach Me, control
    m_ControlHandlers.Add handler, control.name
End Sub

' --- 辅助方法 ---

Public Sub AdjustSize()
    Dim bottomMost As Single
    Dim rightMost As Single
    Dim ctrl As MSForms.control
    
    For Each ctrl In Me.Controls
        If ctrl.top + ctrl.height > bottomMost Then
            bottomMost = ctrl.top + ctrl.height
        End If
        If ctrl.left + ctrl.width > rightMost Then
            rightMost = ctrl.left + ctrl.width
        End If
    Next ctrl
    
    Me.width = rightMost + 20
    Me.height = bottomMost + 40
End Sub

