VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} askdir 
   Caption         =   "springWD"
   ClientHeight    =   3840
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   6015
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
Private WithEvents txt_TM As MSForms.textbox
Attribute txt_TM.VB_VarHelpID = -1
Private colls As New collection
Private class_ctrl As clsCtrls

Private Sub UserForm_Initialize()
    'Call Set_Form
    ' 设置窗体基本属性
    Me.caption = "零件号更新和路径选择"
    Me.width = 300
    Me.height = 200
    Me.StartUpPosition = 1
    Call initFrm
End Sub
Sub initFrm()

    Set chk_TM = Me.Controls.Add("Forms.CheckBox.1", "cjk", True)
        With chk_TM
            .Top = fmargin
            .height = 20
            .name = "chk_TM"
            .caption = "是否更新CATIA零件号时间戳"
            .Left = fmargin
            .width = 250
        End With
        
    Set class_ctrl = New clsCtrls
    
    Set txt_TM = Me.Controls.Add("Forms.TextBox.1", "txt_TM", True)
        With txt_TM
            .Enabled = False
            .Left = fmargin + 20
            .Top = chk_TM.Top + chk_TM.height + itemgap
            .width = 240
            .height = 20
        End With
    
     Set class_ctrl.Txt = txt_TM
    class_ctrl.ohint = "日=d或day，时=hour或h，分=min或i，默认日"
    
    txt_TM.Text = class_ctrl.ohint
    txt_TM.ForeColor = &H808080
    
    colls.Add class_ctrl
    Set chk_Path = Me.Controls.Add("Forms.CheckBox.1")
        With chk_Path
            .name = "chk_Path"
            .caption = "是否导出到当前路径"
            .Left = fmargin
            .Top = txt_TM.Top + txt_TM.height + itemgap
            .width = 150
            .height = 20
        End With
   Set label_log = Me.Controls.Add("Forms.Label.1", "label_log", True)
        With label_log
            .name = "label_log"
            .caption = "设计演进log"
            .Enabled = True
            .Left = fmargin + 4
            .width = 40
            .Top = chk_Path.Top + chk_Path.height + itemgap
            .height = 20
            .BorderStyle = fmBorderStyleSingle
        End With
  
   
   Set txt_log = Me.Controls.Add("Forms.TextBox.1", "txt_log", True)
        With txt_log
            .Enabled = True
            .Left = fmargin + label_log.width + 12
            .Top = chk_Path.Top + chk_Path.height + itemgap
            .width = 200
            .height = 40
        End With
            
    Set cmdOK = Me.Controls.Add("Forms.CommandButton.1")
        With cmdOK
            .name = "cmdOK"
            .caption = "确定"
            .Left = (Me.InsideWidth - (80 * 2) - itemgap) / 2
            .Top = txt_log.Top + txt_log.height + 20
            .width = 60
            .height = 25
        End With
    
    Set cmdCancel = Me.Controls.Add("Forms.CommandButton.1")
        With cmdCancel
            .name = "cmdCancel"
            .caption = "取消"
            .Left = cmdOK.Left + cmdOK.width + itemgap
            .Top = cmdOK.Top
            .width = 60
            .height = 25
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
        txt_TM.ForeColor = &H808080
  End If
End Sub

Private Sub cmdOK_Click()
    Dim UpdateTimestamp As Boolean
    Dim ExportToCurrentPath As Boolean
    UpdateTimestamp = chk_TM.value
    ExportToCurrentPath = chk_Path.value
    Dim tdy(2)
        tdy(0) = 0   ' 0 不更新时间戳， 1更新时间戳
        tdy(2) = 0   ' 1: 导出到当前文档所在路径 0： 导出到被选择的路径
        tdy(1) = ""   ' 0 不更新时间戳， 1更新时间戳
    If UpdateTimestamp Then
        tdy(0) = 1
        tdy(1) = txt_TM.Text
    End If
    
    If ExportToCurrentPath Then
        tdy(2) = 1
    End If
    dt_pth_ctrl = tdy '                                        Debug.Print dt_pth_ctrl(0), dt_pth_ctrl(1), dt_pth_ctrl(2)
        Unload Me
End Sub

Private Sub cmdCancel_Click()
    dt_pth_ctrl = Array(0, 0, "")
    Unload Me
End Sub

Private Sub Set_Form(ByVal MPgs As MultiPage, ByVal Cap As String)
    With Me
        Dim requiredInsideHeight
        requiredInsideHeight = MPgs.Top + MPgs.height + ADJUST_F_H + lb_H  '+ FrmMargin(2)
        .height = requiredInsideHeight + (Me.height - Me.InsideHeight)
        .width = MPgs.width + ADJUST_F_W + 2 * FrmMargin(2)
        .caption = Cap
    End With
End Sub



