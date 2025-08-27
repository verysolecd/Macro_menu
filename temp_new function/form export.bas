'VERSION 1.0 CLASS
'Begin
'  MultiUse = -1  'True
'End
'Attribute VB_Name = "CatiaUpdateForm"
'Attribute VB_GlobalNameSpace = False
'Attribute VB_Creatable = False
'Attribute VB_PredeclaredId = True
'Attribute VB_Exposed = False
'Public CATIA As Application
'Public chk_TM
'Public txt_TM As MSForms.TextBox
'Public chk_Path As MSForms.CheckBox
'Public cmdOK As MSForms.CommandButton
'Public cmdCancel As MSForms.CommandButton

Private Const itemgap = 10
Private Const fmargin = 10
Private Const bmargin = 10

' Use WithEvents to handle events for dynamically created controls.
' Declare each variable on its own line to ensure correct typing.
Private WithEvents chk_TM As MSForms.CheckBox
Private WithEvents cmdOK As MSForms.CommandButton
Private WithEvents cmdCancel As MSForms.CommandButton

Public chk_Path As MSForms.CheckBox
Public chk_TM As MSForms.CheckBox
Public txt_TM As MSForms.TextBox
Public date_path
Public dateFMT

Private Sub UserForm_Initialize()
    ' 设置窗体基本属性
    Me.Caption = "零件号更新和路径选择"
    Me.Width = 300
    Me.Height = 300
    Me.StartUpPosition = 1
   
    Set chk_TM = Me.Controls.Add("Forms.CheckBox.1", "cjk", True)
    With chk_TM
        .Name = "chk_TM"
        .Caption = "是否更新CATIA零件号时间戳"
        .Left = fmargin
        .Top = fmargin + itemgap
        .Width = 250
        .Height = 20
    End With

    
    Set txt_TM = Me.Controls.Add("Forms.TextBox.1", "cjk", True)
    With txt_TM
        .Name = "txt_TM"
        .Text = "请输入日期格式： 日=d或day，时=hour或h，分=min或i，默认为日"
        .Enabled = False
        .Left = fmargin + 20
        .Top = chk_TM.Top + chk_TM.Height + itemgap
        .Width = 240
        .Height = 20
    End With
    
    Set chk_Path = Me.Controls.Add("Forms.CheckBox.1")
    With chk_Path
        .Name = "chk_Path"
        .Caption = "是否导出到当前路径"
        .Left = fmargin
        .Top = txt_TM.Top + txt_TM.Height + itemgap * 2
        .Width = 150
        .Height = 20
    End With
    
    Set cmdOK = Me.Controls.Add("Forms.CommandButton.1")
    With cmdOK
        .Name = "cmdOK"
        .Caption = "确定"
        .Left = (Me.InsideWidth - (80 * 2) - itemgap) / 2
        .Top = Me.InsideHeight - 30 - bmargin
        .Width = 80
        .Height = 30
    End With
    
    Set cmdCancel = Me.Controls.Add("Forms.CommandButton.1")
    With cmdCancel
        .Name = "cmdCancel"
        .Caption = "取消" 
        .Left = cmdOK.Left + cmdOK.Width + itemgap
        .Top = cmdOK.Top
        .Width = 80
        .Height = 30
        .Cancel = True ' Allows Esc key to trigger this button
    End With
End Sub

' 事件处理程序必须使用正确的命名约定
Private Sub chk_TM_Click()
    txt_TM.Enabled = chk_TM.Value
End Sub

Private Sub cmdOK_Click()
    Dim UpdateTimestamp As Boolean
    Dim ExportToCurrentPath As Boolean
    UpdateTimestamp = chk_TM.Value
    ExportToCurrentPath = chk_Path.Value
    Set date_path = Array(UpdateTimestamp, Trim(txt_TM.Text), ExportToCurrentPath)
    Debug.Print date_path(0)
    MsgBox date_path(0)
    MsgBox date_path(1)
    MsgBox date_path(2)
    Unload Me
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

' 实际的CATIA操作函数
Private Sub ProcessUpdates(UpdateTimestamp As Boolean, ExportToCurrentPath As Boolean, DateFormat As String)
    ' This function is currently not used.
    MsgBox "发生错误: " & Err.Description, vbCritical
End Sub
