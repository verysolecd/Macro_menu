VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} wd 
   Caption         =   "UserForm1"
   ClientHeight    =   3795
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4965
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
Private Const FrmTitle As String = "我请问你？？"
Private Const Frm_WIDTH As Integer = 300 ' 窗体固定宽度
Private Const Frm_LH_gap As Integer = 12 ' 所有控件左对齐的左边距
Private Const itemgap As Integer = 1
' 控件默认尺寸
Private Const cls_H As Integer = 20 ' 高度
Private Const cls_W As Integer = 200 ' 宽度
Private Const Btn_W As Integer = 75 ' 按钮宽度
Private Const BTN_H As Integer = 25 ' 按钮高度
Private Const Text_W As Integer = 250 ' 输入框宽度（=窗体宽-2*左边距）
Private Const cls_frontsize = 10

Private Const INPUT_HEIGHT_SINGLE As Integer = 20 ' 单行输入框高度
Private Const INPUT_HEIGHT_MULTI As Integer = 60 ' 多行输入框高度
Private Const OPTION_HEIGHT As Integer = 18 ' 单选/复选框高度
Private bttop As Integer
Private currTop As Long
Private WithEvents ctr As control
Attribute ctr.VB_VarHelpID = -1

' 样式常量（保持美观）
Private Const FONT_NAME As String = "Thoma"
Private Const FONT_SIZE As Integer = 10
Private Const Frm_BACKCOLOR As Long = &H8000000F ' 浅灰背景
Private Const BTN_BACKCOLOR As Long = &H8000000E ' 按钮灰蓝


Private WithEvents mBtn As MSForms.CommandButton
Attribute mBtn.VB_VarHelpID = -1
Private WithEvents mchk As MSForms.CheckBox
Attribute mchk.VB_VarHelpID = -1
Private lst


Sub setFrm(inf)
  With Me
        .Caption = FrmTitle
        .Width = Frm_WIDTH
        .BackColor = Frm_BACKCOLOR
        .Font.Name = FONT_NAME
        .Font.Size = 10
        .StartUpPosition = 2 ' 居中
        .Height = 280
    End With
    bttop = 0
   Set lst = inf
    currTop = 4
'    wd.Show modeless
For Each cfg In lst
    Set ctr = Me.controls.Add(cfg("Type"), cfg("Name"), True)
             With ctr
                    .Font.Size = cls_frontsize
                    .Name = cfg("Name")
                    .Left = Frm_LH_gap
                    .Width = cls_W
              Select Case cfg("Type")
                    Case "Forms.CommandButton.1"
                             If bttop = 0 Then
                                  bttop = currTop
                                  .top = bttop  '98
                                  thisleft = .Left + Btn_W + itemgap
                                  Debug.Print .top
                             Else
                                  .top = bttop
                  
                                  Debug.Print "第二按钮高度" & .top
                                 .Left = thisleft
                             End If
                             .Width = Btn_W
                        Case Else
                            .top = currTop
                     End Select
                    .Height = cls_H
                     currTop = .top + .Height + itemgap
                      Debug.Print currTop
                   If cfg("Type") <> "Forms.TextBox.1" Then
                       .Caption = cfg("Caption")
                     Else
                       .Text = cfg("Caption")
                        .Width = Me.Width - 3 * Frm_LH_gap
                    End If
            End With
        Next
        
        Me.Height = (lst.count + 1) * (cls_H + itemgap)
End Sub
Private Sub mchk_Click()

Me.controls.item ("lst.")

  txt_TM.Enabled = chk_TM.value

End Sub

Private Sub UserForm_Click()

End Sub
