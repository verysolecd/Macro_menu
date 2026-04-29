VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} askdir 
   Caption         =   "请确定时间格式和路径"
   ClientHeight    =   3135
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   7845
   OleObjectBlob   =   "askdir.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "askdir"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'VERSION 1.0 CLASS
'Begin
'  MultiUse = -1  'True
'End

Private Const itemgap = 3
Private Const fmargin = 6
Private Const bmargin = 6
' Use WithEvents to handle events for dynamically created controls.
Private WithEvents chk_TM As MSForms.CheckBox
Attribute chk_TM.VB_VarHelpID = -1
Private WithEvents chk_Path As MSForms.CheckBox
Attribute chk_Path.VB_VarHelpID = -1
Private WithEvents cmdOK As MSForms.CommandButton
Attribute cmdOK.VB_VarHelpID = -1
Private WithEvents cmdCancel As MSForms.CommandButton
Attribute cmdCancel.VB_VarHelpID = -1
Private WithEvents txt_TM As MSForms.TextBox
Attribute txt_TM.VB_VarHelpID = -1
Private colls As New collection
Private class_ctrl As clsCtrls

Private Sub UserForm_Initialize()
    ' 设置窗体基本属性
    Me.Caption = "零件号更新和路径选择"
    Me.Width = 300
    Me.Height = 130
    Me.StartUpPosition = 1
   
    Set chk_TM = Me.Controls.Add("Forms.CheckBox.1", "cjk", True)
        With chk_TM
            .Name = "chk_TM"
            .Caption = "是否更新CATIA零件号时间戳"
            .Left = fmargin
            .Top = fmargin
            .Width = 250
            .Height = 20
        End With
        
    Set class_ctrl = New clsCtrls
    
    Set txt_TM = Me.Controls.Add("Forms.TextBox.1", "txt_TM", True)
    
        With txt_TM
            .Enabled = False
            .Left = fmargin + 20
            .Top = chk_TM.Top + chk_TM.Height + itemgap
            .Width = 240
            .Height = 20
        End With
    
     Set class_ctrl.txt = txt_TM
    class_ctrl.ohint = "请输入日期格式，日=d或day，时=hour或h，分=min或i，默认为日"
    
    txt_TM.Text = class_ctrl.ohint
    txt_TM.ForeColor = &H808080
    
    colls.Add class_ctrl
    
    Set chk_Path = Me.Controls.Add("Forms.CheckBox.1")
        With chk_Path
            .Name = "chk_Path"
            .Caption = "是否导出到当前路径"
            .Left = fmargin
            .Top = txt_TM.Top + txt_TM.Height + itemgap
            .Width = 150
            .Height = 20
        End With
        
    Set cmdOK = Me.Controls.Add("Forms.CommandButton.1")
        With cmdOK
            .Name = "cmdOK"
            .Caption = "确定"
            .Left = (Me.InsideWidth - (80 * 2) - itemgap) / 2
            .Top = chk_Path.Top + 20
            .Width = 60
            .Height = 25
        End With
    
    Set cmdCancel = Me.Controls.Add("Forms.CommandButton.1")
        With cmdCancel
            .Name = "cmdCancel"
            .Caption = "取消"
            .Left = cmdOK.Left + cmdOK.Width + itemgap
            .Top = cmdOK.Top
            .Width = 60
            .Height = 25
            .Cancel = True ' Allows Esc key to trigger this button
        End With
End Sub

Private Sub chk_TM_Click()
    txt_TM.Enabled = chk_TM.Value
End Sub

Private Sub txt_TM_gotfocus()
    If txt_TM.Text = usrTXT Then
        txt_TM.Text = " "
        txt_TM.ForeColor = &H80
    End If
End Sub

Private Sub txt_TM_Lostfocus()
    If txt_TM.Text = "" Then
        txt_TM.Text = usrTXT
        txt_TM.ForeColor = &H808080
  End If
End Sub

Private Sub cmdOK_Click()
    Dim UpdateTimestamp As Boolean
    Dim ExportToCurrentPath As Boolean
    UpdateTimestamp = chk_TM.Value
    ExportToCurrentPath = chk_Path.Value
    Dim tdy(2)
        tdy(0) = 0
        tdy(2) = 0
        tdy(1) = ""
    If UpdateTimestamp Then
        tdy(0) = 1
        tdy(1) = txt_TM.Text
    End If
    
    If ExportToCurrentPath Then
        tdy(2) = 1
    End If

    dt_pth_ctrl = tdy
'                                        Debug.Print dt_pth_ctrl(0), dt_pth_ctrl(1), dt_pth_ctrl(2)
        Unload Me
End Sub

Private Sub cmdCancel_Click()

    dt_pth_ctrl = Array(0, 0, "")
    Unload Me
End Sub



