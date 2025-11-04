Attribute VB_Name = "springWD"
'Attribute VB_Name = "springWD"
' 自定义弹窗模块，使用现有的askdir窗体并动态添加控件
Option Explicit

' 弹窗结果枚举
Public Enum PopupResult
    prOK = 1
    prCancel = 2
End Enum

' 控件信息结构
Private Type ControlInfo
    Type As Integer ' 1=Checkbox, 2=TextBox
    name As String
    caption As String
    X As Integer
    y As Integer
    width As Integer
    height As Integer
    value As Variant
End Type

' 模块级变量
Private Controls() As ControlInfo
Private ControlCount As Integer
Private PopupTitle As String
Private result As PopupResult
Private Values As Object

' 创建新弹窗
' 参数:
'   title - 弹窗标题
Public Sub CreatePopup(Optional title As String = "弹窗")
    ControlCount = 0
    ReDim Controls(0)
    PopupTitle = title
    result = prCancel
    Set Values = CreateObject("Scripting.Dictionary")
End Sub

' 添加复选框控件
' 参数:
'   name - 控件名称(用于后续获取值)
'   caption - 显示文本
'   x, y - 位置坐标
'   width, height - 控件尺寸
'   defaultValue - 默认选中状态(True/False)
Public Sub AddCheckbox(name As String, caption As String, X As Integer, y As Integer, width As Integer, height As Integer, Optional defaultValue As Boolean = False)
    ControlCount = ControlCount + 1
    ReDim Preserve Controls(ControlCount)
    
    With Controls(ControlCount)
        .Type = 1
        .name = name
        .caption = caption
        .X = X
        .y = y
        .width = width
        .height = height
        .value = defaultValue
    End With
End Sub

' 添加文本框控件
' 参数:
'   name - 控件名称(用于后续获取值)
'   caption - 标签文本
'   x, y - 位置坐标
'   width, height - 控件尺寸
'   defaultValue - 默认文本
Public Sub AddTextBox(name As String, caption As String, X As Integer, y As Integer, width As Integer, height As Integer, Optional defaultValue As String = "")
    ControlCount = ControlCount + 1
    ReDim Preserve Controls(ControlCount)
    
    With Controls(ControlCount)
        .Type = 2
        .name = name
        .caption = caption
        .X = X
        .y = y
        .width = width
        .height = height
        .value = defaultValue
    End With
End Sub

' 显示弹窗并返回结果
Public Function ShowPopup() As PopupResult
    Dim askdirForm As Object
    Dim i As Integer
    Dim Ctrl As ControlInfo
    Dim maxWidth As Integer
    Dim maxHeight As Integer
    Dim gap As Integer
    
    ' 使用现有的askdir窗体
    On Error Resume Next
    Set askdirForm = New askdir
    On Error GoTo 0
    
    If askdirForm Is Nothing Then
        MsgBox "无法创建askdir窗体", vbExclamation, "错误"
        ShowPopup = prCancel
        Exit Function
    End If
    
    ' 设置窗体标题
    askdirForm.caption = PopupTitle
    

    
    ' 计算窗体所需尺寸
    maxWidth = 0
    maxHeight = 0
    gap = 20 ' 控件间的间隙
    
    ' 先找出控件的最大坐标和尺寸
    For i = 1 To ControlCount
        Ctrl = Controls(i)
        If Ctrl.X + Ctrl.width > maxWidth Then
            maxWidth = Ctrl.X + Ctrl.width
        End If
        If Ctrl.y + Ctrl.height > maxHeight Then
            maxHeight = Ctrl.y + Ctrl.height
        End If
    Next i
    
    ' 设置窗体尺寸（加上边距）
    askdirForm.width = maxWidth + gap * 4
    askdirForm.height = maxHeight + gap * 3 ' 多加一些高度用于按钮
    
    ' 添加控件到窗体
    For i = 1 To ControlCount
        Ctrl = Controls(i)
        Select Case Ctrl.Type
            Case 1 ' Checkbox
                With askdirForm.Controls.Add("Forms.CheckBox.1")
                    .name = Ctrl.name
                    .caption = Ctrl.caption
                    .Left = Ctrl.X
                    .Top = Ctrl.y
                    .width = Ctrl.width
                    .height = Ctrl.height
                    .value = Ctrl.value
                End With
                
            Case 2 ' TextBox
                ' 添加标签
                With askdirForm.Controls.Add("Forms.Label.1")
                    .name = "lbl_" & Ctrl.name
                    .caption = Ctrl.caption
                    .Left = Ctrl.X
                    .Top = Ctrl.y + 3 ' 微调位置使与文本框对齐
                    .width = Ctrl.width
                    .height = Ctrl.height
                End With
                
                ' 添加文本框
                With askdirForm.Controls.Add("Forms.TextBox.1")
                    .name = Ctrl.name
                    .Left = Ctrl.X + Ctrl.width + 10
                    .Top = Ctrl.y
                    .width = Ctrl.width * 2
                    .height = Ctrl.height
                    .Text = Ctrl.value
                End With
        End Select
    Next i
    
    ' 添加确定和取消按钮
    ' 确定按钮
    With askdirForm.Controls.Add("Forms.CommandButton.1")
        .name = "cmdOK"
        .caption = "确定"
        .Left = askdirForm.width - 180
        .Top = askdirForm.height - 60
        .width = 75
        .height = 25
        .Default = True
    End With
    
    ' 取消按钮
    With askdirForm.Controls.Add("Forms.CommandButton.1")
        .name = "cmdCancel"
        .caption = "取消"
        .Left = askdirForm.width - 90
        .Top = askdirForm.height - 60
        .width = 75
        .height = 25
    End With
    
    ' 显示模态窗体
    askdirForm.Show vbModal
    
    ' 收集结果
    If result = prOK Then
        Values.RemoveAll
        For i = 1 To ControlCount
            Ctrl = Controls(i)
            On Error Resume Next
            Select Case Ctrl.Type
                Case 1 ' Checkbox
                    Values.Add Ctrl.name, askdirForm.Controls(Ctrl.name).value
                Case 2 ' TextBox
                    Values.Add Ctrl.name, askdirForm.Controls(Ctrl.name).Text
            End Select
            On Error GoTo 0
        Next i
    End If
    
    ' 清理窗体引用
    Set askdirForm = Nothing
    
    ShowPopup = result
End Function

' 用于处理OK按钮点击事件的函数
' 需要在askdir窗体的cmdOK按钮Click事件中调用
Public Sub HandleOK()
    result = prOK
End Sub

' 用于处理Cancel按钮点击事件的函数
' 需要在askdir窗体的cmdCancel按钮Click事件中调用
Public Sub HandleCancel()
    result = prCancel
End Sub

' 获取控件的值
' 参数: name - 控件名称
' 返回值: 控件的值
Public Function GetValue(name As String) As Variant
    On Error Resume Next
    GetValue = Values(name)
    On Error GoTo 0
End Function
