VERSION 5.00
Begin VB.Form DatePathForm 
    Caption         =   "日期和路径设置"
    ClientHeight    =   2100
    ClientLeft      =   45
    ClientTop       =   330
    ClientWidth     =   4680
    Icon            =   "DatePathForm.frx":0000
    LinkTopic       =   "Form1"
    ScaleHeight     =   2100
    ScaleWidth      =   4680
    StartUpPosition =   1  '所有者中心
    Begin VB.CommandButton cmdExecute 
        Caption         =   "执行操作"
        Height          =   495
        Left            =   1440
        TabIndex        =   3
        Top             =   1440
        Width           =   1695
    End
    Begin VB.CheckBox chkPath 
        Caption         =   "更新路径"
        Height          =   375
        Left            =   720
        TabIndex        =   2
        Top             =   960
        Width           =   2895
    End
    Begin VB.TextBox txtDate 
        Enabled         =   0   'False
        Height          =   375
        Left            =   1800
        TabIndex        =   1
        Text            =   "YYYY-MM-DD"
        Top             =   480
        Width           =   2175
    End
    Begin VB.CheckBox chkDate 
        Caption         =   "更新日期"
        Height          =   375
        Left            =   720
        TabIndex        =   0
        Top             =   480
        Width           =   1095
    End
    Begin VB.Label Label1 
        Caption         =   "请选择需要执行的操作："
        Height          =   375
        Left            =   720
        TabIndex        =   4
        Top             =   120
        Width           =   2895
    End
End
Attribute VB_Name = "DatePathForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

' 执行按钮点击事件
Private Sub cmdExecute_Click()
    Dim updateDate As Boolean
    Dim updatePath As Boolean
    Dim dateValue As String
    
    ' 获取复选框状态
    updateDate = chkDate.Value
    updatePath = chkPath.Value
    
    ' 验证日期选择
    If updateDate Then
        ' 检查日期文本框是否有值
        If Trim(txtDate.Text) = "" Or txtDate.Text = "YYYY-MM-DD" Then
            MsgBox "请在日期文本框中输入有效的日期值", vbExclamation, "输入错误"
            txtDate.SetFocus
            Exit Sub
        End If
        
        ' 简单的日期格式验证
        If Not IsDate(txtDate.Text) Then
            MsgBox "请输入有效的日期格式（如：2023-12-31）", vbExclamation, "格式错误"
            txtDate.SetFocus
            Exit Sub
        End If
        
        ' 保存日期值
        dateValue = txtDate.Text
    End If
    
    ' 执行操作
    ExecuteOperations updateDate, updatePath, dateValue
End Sub

' 执行具体操作的过程
Private Sub ExecuteOperations(needUpdateDate As Boolean, needUpdatePath As Boolean, newDate As String)
    Dim resultMsg As String
    resultMsg = "操作已完成：" & vbCrLf & vbCrLf
    
    ' 处理日期更新
    If needUpdateDate Then
        resultMsg = resultMsg & "? 已更新日期为: " & Format(newDate, "yyyy年mm月dd日") & vbCrLf
        ' 在这里添加实际更新日期的代码
        ' 例如: UpdateSystemDate newDate
    Else
        resultMsg = resultMsg & "? 未更新日期" & vbCrLf
    End If
    
    ' 处理路径更新
    If needUpdatePath Then
        resultMsg = resultMsg & "? 已更新路径" & vbCrLf
        ' 在这里添加实际更新路径的代码
        ' 例如: UpdateFilePath
    Else
        resultMsg = resultMsg & "? 未更新路径" & vbCrLf
    End If
    
    ' 显示操作结果
    MsgBox resultMsg, vbInformation, "操作结果"
End Sub

' 日期复选框状态变化事件
Private Sub chkDate_Click()
    ' 当勾选日期复选框时启用文本框，否则禁用
    txtDate.Enabled = chkDate.Value
    
    ' 如果是勾选状态且文本框还是默认提示文字，则清空
    If chkDate.Value And txtDate.Text = "YYYY-MM-DD" Then
        txtDate.Text = ""
    ElseIf Not chkDate.Value And txtDate.Text = "" Then
        ' 如果取消勾选且文本框为空，则恢复提示文字
        txtDate.Text = "YYYY-MM-DD"
    End If
End Sub

' 窗体初始化事件
Private Sub UserForm_Initialize()
    ' 设置初始状态
    txtDate.Enabled = False
    chkDate.Value = False
    chkPath.Value = False
End Sub
