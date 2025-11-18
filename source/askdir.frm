VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} askdir 
   Caption         =   "springWD"
   ClientHeight    =   3840
   ClientLeft      =   120
   ClientTop       =   470
   ClientWidth     =   6020
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
Private Const itemgap = 2
Private Const fmargin = 6
Private Const bmargin = 6
Private Const txtWidth = 236

' Use WithEvents to handle events for dynamically created controls.
Private WithEvents chk_TM As MSForms.CheckBox
Attribute chk_TM.VB_VarHelpID = -1
Private WithEvents chk_Path As MSForms.CheckBox
Attribute chk_Path.VB_VarHelpID = -1
Private WithEvents cmdOK As MSForms.CommandButton
Attribute cmdOK.VB_VarHelpID = -1
Private WithEvents cmdCancel As MSForms.CommandButton
Attribute cmdCancel.VB_VarHelpID = -1
Private WithEvents txt_TM As MSForms.textbox
Attribute txt_TM.VB_VarHelpID = -1
Private WithEvents txt_log As MSForms.textbox
Attribute txt_log.VB_VarHelpID = -1
Private WithEvents chk_log As MSForms.CheckBox
Attribute chk_log.VB_VarHelpID = -1


Private colls As New collection
Private class_ctrl As clsCtrls
Private Sub UserForm_Initialize()
    'Call Set_Form
    ' 设置窗体基本属性
    Me.Caption = "零件号更新和路径选择"
    Me.Width = 280
    Me.Height = 200
    Me.StartUpPosition = 1
    Call initFrm
End Sub
Sub initFrm()

Dim thistop
    Set chk_Path = Me.controls.Add("Forms.CheckBox.1")
           With chk_Path
               .Name = "chk_Path":    .Caption = "是否导出到当前路径"
               .Left = fmargin:     .Width = 150
               .top = fmargin:     .Height = 20
               thistop = .top + .Height + itemgap
           End With
    Set chk_TM = Me.controls.Add("Forms.CheckBox.1", "chk_TM", True)
        With chk_TM
            .Name = "chk_TM":  .Caption = "是否更新CATIA零件号时间戳"
            .Left = fmargin:   .Width = 240
            .top = thistop:   .Height = 20
             thistop = .top + .Height + itemgap
        End With
    Set class_ctrl = New clsCtrls
    Set txt_TM = Me.controls.Add("Forms.TextBox.1", "txt_TM", True)
        With txt_TM
            .Name = "txt_TM"
            .Enabled = False
            .Left = fmargin + 12: .Width = txtWidth
            .top = thistop: .Height = 20
              thistop = .top + .Height + itemgap
        End With
     Set class_ctrl.Txt = txt_TM
        class_ctrl.ohint = "日=d或day，时=hour或h，分=min或i，默认日"
        txt_TM.Text = class_ctrl.ohint
        txt_TM.ForeColor = &H808080
    colls.Add class_ctrl
    
   Set chk_log = Me.controls.Add("Forms.CheckBox.1", "chk_log", True)
    With chk_log
            .Name = "chk_log": .Caption = "是否更新本次导出日志"
            .top = thistop: .Height = 20
            .Left = fmargin: .Width = 240
            thistop = .top + .Height + itemgap
         End With
   
   Set txt_log = Me.controls.Add("Forms.TextBox.1", "txt_log", True)
        With txt_log
            .Name = "txt_log"
            .Left = fmargin + 12: .Width = txtWidth
            .top = thistop: .Height = 40
            thistop = .top + .Height + itemgap
        End With
'        Debug.Print "log已经创建"
   
    Set cmdOK = Me.controls.Add("Forms.CommandButton.1")
        With cmdOK
            .Name = "cmdOK":    .Caption = "确定"
            .Left = (Me.InsideWidth - (80 * 2) - itemgap) / 2:     .Width = 60
            .top = thistop:     .Height = 24
            thistop = .top + .Height + itemgap
        End With
    Set cmdCancel = Me.controls.Add("Forms.CommandButton.1")
        With cmdCancel
            .Name = "cmdCancel": .Caption = "取消"
            .Left = cmdOK.Left + cmdOK.Width + itemgap: .Width = 60
            .top = cmdOK.top: .Height = 24
            .Cancel = True ' Allows Esc key to trigger this button
        End With
End Sub
Private Sub chk_TM_Click()
    txt_TM.Enabled = chk_TM.value
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
        txt_TM.ForeColor = &H80000012
  End If
End Sub
Private Sub cmdOK_Click()
    Dim UpdateTimestamp As Boolean
    UpdateTimestamp = chk_TM.value
    Dim ExportToCurrentPath As Boolean
    ExportToCurrentPath = chk_Path.value
    
    Dim tdy(3)
        tdy(0) = 0   ' 0 不更新时间戳， 1更新时间戳
        tdy(1) = ""   ' 时间戳格式
        tdy(2) = 0   ' 1: 导出到当前文档所在路径 0： 导出到被选择的路径
        tdy(3) = ""  '导出日志内容，空时不导入
    If UpdateTimestamp Then
        tdy(0) = 1
        tdy(1) = txt_TM.Text
    End If
    If ExportToCurrentPath Then
        tdy(2) = 1
    End If
    Dim log
       If chk_log.value Then
            tdy(3) = txt_log.value
        End If
   
    export_CFG = tdy '   Debug.Print export_CFG(0), export_CFG(1), export_CFG(2)m,export_CFG(3)
        Unload Me
End Sub
Private Sub cmdCancel_Click()
    export_CFG = Array(0, 0, "", "")
    Unload Me
End Sub
Private Sub Set_Form(ByVal MPgs As MultiPage, ByVal Cap As String)
    With Me
        Dim requiredInsideHeight
        requiredInsideHeight = MPgs.top + MPgs.Height + ADJUST_F_H + lb_H  '+ FrmMargin(2)
        .Height = requiredInsideHeight + (Me.Height - Me.InsideHeight)
        .Width = MPgs.Width + ADJUST_F_W + 2 * FrmMargin(2)
        .Caption = Cap
    End With
End Sub
