VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} wd 
   Caption         =   "UserForm1"
   ClientHeight    =   3015
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   OleObjectBlob   =   "wd.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "wd"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

' 模块：modStyle（简化版）
' 布局常量（核心简化点）
Private Const FORM_WIDTH As Integer = 300 ' 窗体固定宽度
Private Const LEFT_MARGIN As Integer = 15 ' 所有控件左对齐的左边距
Private Const CONTROL_SPACING As Integer = 10 ' 控件间垂直间距
Private Const TOP_START As Integer = 15 ' 第一个控件的顶部起始位置
' 控件默认尺寸
Private Const LABEL_HEIGHT As Integer = 15 ' 标签高度
Private Const BTN_WIDTH As Integer = 80 ' 按钮宽度
Private Const BTN_HEIGHT As Integer = 25 ' 按钮高度
Private Const INPUT_WIDTH As Integer = 250 ' 输入框宽度（=窗体宽-2*左边距）
Private Const INPUT_HEIGHT_SINGLE As Integer = 20 ' 单行输入框高度
Private Const INPUT_HEIGHT_MULTI As Integer = 60 ' 多行输入框高度
Private Const OPTION_HEIGHT As Integer = 18 ' 单选/复选框高度
' 样式常量（保持美观）
Private Const FONT_NAME As String = "微软雅黑"
Private Const FONT_SIZE As Integer = 10
Private Const FORM_BACKCOLOR As Long = &H8000000F ' 浅灰背景
Private Const BTN_BACKCOLOR As Long = &H8000000D ' 按钮灰蓝
' 模块：modFormGenerator（简化版）
Public Sub ShowSimpleForm(controls As collection, formTitle As String, callback As String)
    Dim frm As Object
    Dim ctl As Object, ctlConfig As clsControlConfig
    Dim currentTop As Integer ' 当前顶部坐标（累计值）
    ' 1. 创建窗体（固定宽度）
    Set frm = CreateObject("Forms.Form.1")
    With frm
        .Caption = formTitle
        .Width = FORM_WIDTH
        .BackColor = FORM_BACKCOLOR
        .Font.Name = FONT_NAME
        .Font.Size = FONT_SIZE
        .StartUpPosition = 2 ' 居中
    End With
    currentTop = TOP_START ' 从顶部起始位置开始
    ' 2. 循环创建控件（按顺序排列）
    For Each ctlConfig In controls
        ' 创建控件并设置位置和尺寸
        Set ctl = CreateControl(frm, ctlConfig, currentTop)
        ' 更新当前顶部坐标（下一个控件的位置）
        currentTop = currentTop + ctl.Height + CONTROL_SPACING
    Next
    ' 3. 调整窗体高度（最后一个控件底部+底部边距）
    frm.Height = currentTop + TOP_START ' 底部留同样的边距
    ' 4. 绑定事件（按钮点击）
    BindEvents frm, callback
    ' 5. 显示窗体
    frm.Show vbModal
End Sub
' 创建单个控件（使用默认尺寸和固定左对齐）
Private Function CreateControl(frm As Object, cfg As clsControlConfig, top As Integer) As Object
    Dim ctl As Object
    Select Case cfg.ControlType
        Case "Label"
            Set ctl = frm.controls.Add("Forms.Label.1", cfg.Name)
            With ctl
                .Caption = cfg.Caption
                .Height = LABEL_HEIGHT
                .AutoSize = True ' 标签宽度自适应文本
            End With
        Case "CommandButton"
            Set ctl = frm.controls.Add("Forms.CommandButton.1", cfg.Name)
            With ctl
                .Caption = cfg.Caption
                .Width = BTN_WIDTH
                .Height = BTN_HEIGHT
                .BackColor = BTN_BACKCOLOR
            End With
        Case "TextBox"
            Set ctl = frm.controls.Add("Forms.TextBox.1", cfg.Name)
            With ctl
                .Text = cfg.DefaultValue
                .Width = INPUT_WIDTH
                .Height = IIf(cfg.MultiLine, INPUT_HEIGHT_MULTI, INPUT_HEIGHT_SINGLE)
                .MultiLine = cfg.MultiLine
                .BorderStyle = 1
            End With
        Case "OptionButton"
            Set ctl = frm.controls.Add("Forms.OptionButton.1", cfg.Name)
            With ctl
                .Caption = cfg.Caption
                .Height = OPTION_HEIGHT
                .value = (cfg.DefaultValue = cfg.Caption) ' 默认选中项
            End With
        Case "CheckBox"
            Set ctl = frm.controls.Add("Forms.CheckBox.1", cfg.Name)
            With ctl
                .Caption = cfg.Caption
                .Height = OPTION_HEIGHT
                .value = cfg.DefaultValue
            End With
    End Select
    ' 固定左对齐（所有控件Left相同）
    ctl.Left = LEFT_MARGIN
    ctl.top = top ' 使用传入的顶部坐标
    Set CreateControl = ctl
End Function
' 模块：modEvents（简化版）

' 模块：modTest
Sub TestSimpleForm()
    Dim controls As New collection, ctl As clsControlConfig
    ' 1. 标题标签
    Set ctl = New clsControlConfig
    ctl.ControlType = "Label"
    ctl.Name = "lblTitle"
    ctl.Caption = "简单信息采集"
    controls.Add ctl
    ' 2. 单选框
    Set ctl = New clsControlConfig
    ctl.ControlType = "OptionButton"
    ctl.Name = "opt1"
    ctl.Caption = "选项A"
    ctl.DefaultValue = "选项A"
    controls.Add ctl
    ' 3. 输入框
    Set ctl = New clsControlConfig
    ctl.ControlType = "TextBox"
    ctl.Name = "txtInput"
    ctl.Caption = "（无需用到，文本框靠DefaultValue显示提示）"
    ctl.DefaultValue = "请输入内容"
    controls.Add ctl
    ' 4. 按钮
    Set ctl = New clsControlConfig
    ctl.ControlType = "CommandButton"
    ctl.Name = "btnOK"
    ctl.Caption = "确定"
    controls.Add ctl
    ' 显示窗体
    ShowSimpleForm controls, "简化版动态窗体", "HandleResult"
End Sub
Sub HandleResult(result As Scripting.Dictionary)
    MsgBox "用户输入：" & result("txtInput") & vbCrLf & "选中项：" & result("SelectedOption")
End Sub
Private Sub UserForm_Click()
End Sub
