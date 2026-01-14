VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} springWD 
   Caption         =   "UserForm1"
   ClientHeight    =   915
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   1800
   OleObjectBlob   =   "springWD.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "springWD"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

' 模块：modStyle（简化版）
' 布局常量（核心简化点）

Private Const Frm_LH_gap As Integer = 6 ' 所有控件左对齐的左边距
Private Const ItmGap As Integer = 0.8
' 控件默认尺寸
Private Const cls_H As Integer = 19.5 ' 高度
Private Const cls_W As Integer = 220 ' 宽度
Private Const Btn_W As Integer = 60 ' 按钮宽度
Private Const Text_W As Integer = 280
Private Const cls_frontsize = 11
' 样式常量（保持美观）
Private Const FONT_NAME As String = "Thoma"
Private Const Frm_color As Long = &H8000000F ' 浅灰背景
Private Const BTN_color As Long = &H8000000E ' 按钮灰蓝
Private lst, cfg, ctr
Private reqHeight, reqWidth, curRH, curBtm
Private BtnTop, BtnLeft, currTop
Option Explicit
Sub setFrm(ttl, inf)
    BtnTop = 0
    currTop = 0
    BtnLeft = Frm_LH_gap
    reqWidth = 0
    reqHeight = 0
    Set lst = inf
                        
     Dim Textlst
     Set Textlst = KCL.InitLst
On Error Resume Next
    For Each cfg In lst
        Set ctr = Me.Controls.Add(cfg("Type"), cfg("Name"), True)
        With ctr
            .Name = cfg("Name"): .Caption = cfg("Caption")
            .Left = Frm_LH_gap: .Width = cls_W
            .Font.Name = FONT_NAME: .Font.Size = cls_frontsize
            .Height = cls_H
            
            '以下设置ctr Top并重置必要的left
                Select Case cfg("Type")
                    Case "Forms.CommandButton.1"
                        If BtnTop = 0 Then BtnTop = IIf(currTop = 0, 15, currTop + 3 * ItmGap) ' 如果是第一个按钮
                        .Top = BtnTop
                        .Left = BtnLeft: .Width = Btn_W
                        BtnLeft = .Left + .Width + 1.5 * ItmGap ' 计算下一个按钮的左边距
                        .BackColor = BTN_color
                         currTop = BtnTop
                         .Font.Size = 10
                         
                    Case "Forms.TextBox.1"
                             .AutoSize = False
                            Textlst.Add ctr
                            .Top = currTop
                            .Width = Text_W
                            .Height = 2 * cls_H
                            .Text = cfg("Caption")
                    Case Else
                        .AutoSize = True
                        .Caption = cfg("Caption")
                        .Top = currTop
                End Select
            currTop = .Top + .Height
        End With
    Next
On Error GoTo 0
    For Each ctr In Me.Controls
      If ctr.Visible Then
            With ctr
                curRH = .Left + .Width
                curBtm = .Top + .Height
            End With
        End If
        reqWidth = IIf(curRH > reqWidth, curRH, reqWidth)
        reqHeight = IIf(curBtm > reqHeight, curBtm, reqHeight)
    Next
    With Me
        .Caption = ttl
        .BackColor = Frm_color
        .Font.Name = FONT_NAME: .Font.Size = cls_frontsize
        .StartUpPosition = 2 ' 居中
        .Height = reqHeight + (Me.Height - Me.InsideHeight) + 6 * ItmGap
        .Width = reqWidth + (.Width - .InsideWidth) + Frm_LH_gap
    End With
    
    Dim txt
    For Each txt In Textlst
        txt.Width = Me.Width - 4 * Frm_LH_gap
    Next
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If CloseMode = 0 Then
        Me.Tag = "UserClosed"
        Cancel = 1
        Me.Hide
    End If
End Sub

Private Sub UserForm_Click()

End Sub

