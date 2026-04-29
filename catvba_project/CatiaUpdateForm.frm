VERSION 5.00
Begin VB.Form CatiaUpdateForm 
    Caption         =   "CATIA更新设置"
    ClientHeight    =   3000
    ClientLeft      =   45
    ClientTop       =   330
    ClientWidth     =   4800
    ControlBox      =   -1  'True
    Icon            =   "CatiaUpdateForm.frx":0000
    LinkTopic       =   "Form1"
    MaxButton       =   0   'False
    MinButton       =   0   'False
    ScaleHeight     =   3000
    ScaleWidth      =   4800
    ShowInTaskbar   =   0   'False
    StartUpPosition =   1  '所有者中心
    Begin VB.CommandButton cmdCancel 
        Caption         =   "取消"
        Height          =   405
        Left            =   2880
        TabIndex        =   4
        Top             =   2340
        Width           =   1215
    End
    Begin VB.CommandButton cmdOK 
        Caption         =   "确定"
        Height          =   405
        Left            =   960
        TabIndex        =   3
        Top             =   2340
        Width           =   1215
    End
    Begin VB.CheckBox chkExportPath 
        Caption         =   "导出到当前路径"
        Height          =   375
        Left            =   720
        TabIndex        =   2
        Top             =   1560
        Width           =   3375
    End
    Begin VB.TextBox txtDateFormat 
        Enabled         =   0   'False
        Height          =   375
        Left            =   720
        TabIndex        =   1
        Text            =   "yyyyMMddHHmmss"
        Top             =   960
        Width           =   3375
    End
    Begin VB.CheckBox chkUpdateTimestamp 
        Caption         =   "是否更新catia零件号时间戳"
        Height          =   375
        Left            =   720
        TabIndex        =   0
        Top             =   480
        Width           =   3375
    End
    Begin VB.Label Label1 
        Caption         =   "时间格式 (例如: yyyyMMddHHmmss):"
        Height          =   255
        Left            =   720
        TabIndex        =   5
        Top             =   720
        Width           =   3375
    End
End
Attribute VB_Name = "CatiaUpdateForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

' 确定按钮点击事件
Private Sub cmdOK_Click()
    Dim updateTimestamp As Boolean
    Dim exportToCurrentPath As Boolean
    Dim dateFormat As String
    
    ' 获取复选框状态
    updateTimestamp = chkUpdateTimestamp.Value
    exportToCurrentPath = chkExportPath.Value
    
    ' 获取时间格式（如果勾选了更新时间戳）
    If updateTimestamp Then
        dateFormat = Trim(txtDateFormat.Text)
        ' 简单验证时间格式不为空
        If dateFormat = "" Then
            MsgBox "请输入时间格式", vbExclamation, "提示"
            txtDateFormat.SetFocus
            Exit Sub
        End If
    End If
    
    ' 调用处理函数，传递参数
    ProcessUpdates updateTimestamp, exportToCurrentPath, dateFormat
    
    ' 关闭窗体
    Unload Me
End Sub

' 取消按钮点击事件
Private Sub cmdCancel_Click()
    ' 直接关闭窗体，不执行任何操作
    Unload Me
End Sub

' 时间戳复选框状态变化事件
Private Sub chkUpdateTimestamp_Click()
    ' 勾选时启用时间格式文本框，否则禁用
    txtDateFormat.Enabled = chkUpdateTimestamp.Value
End Sub

' 处理更新的函数（这里是框架，你需要根据实际需求实现具体逻辑）
Private Sub ProcessUpdates(UpdateTimestamp As Boolean, ExportToCurrentPath As Boolean, DateFormat As String)
    ' 这里只是示例消息，实际使用时替换为你的函数调用
    Dim msg As String
    msg = "执行结果:" & vbCrLf & vbCrLf
    
    ' 根据选择调用相应的函数
    If UpdateTimestamp Then
        msg = msg & "1. 将更新CATIA零件号时间戳" & vbCrLf
        msg = msg & "   时间格式: " & DateFormat & vbCrLf & vbCrLf
        ' 实际调用更新时间戳的函数
        ' UpdateCatiaTimestamp(DateFormat)
    Else
        msg = msg & "1. 不更新CATIA零件号时间戳" & vbCrLf & vbCrLf
    End If
    
    If ExportToCurrentPath Then
        msg = msg & "2. 将导出到当前路径" & vbCrLf
        ' 实际调用更新路径的函数
        ' ExportToPath(True)
    Else
        msg = msg & "2. 不导出到当前路径" & vbCrLf
        ' 实际调用使用默认路径的函数
        ' ExportToPath(False)
    End If
    
    MsgBox msg, vbInformation, "操作确认"
End Sub

' 窗体初始化事件
Private Sub UserForm_Initialize()
    ' 设置初始状态
    chkUpdateTimestamp.Value = False
    chkExportPath.Value = False
    txtDateFormat.Enabled = False
    txtDateFormat.Text = "yyyyMMddHHmmss" ' 默认时间格式
End Sub
