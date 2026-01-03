VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} WD 
   Caption         =   "UserForm1"
   ClientHeight    =   6780
   ClientLeft      =   120
   ClientTop       =   470
   ClientWidth     =   9860.001
   OleObjectBlob   =   "WD.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "WD"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' 模块：modStyle（简化版）
' 布局常量（核心简化点）

Private Const Frm_WIDTH As Integer = 300 ' 窗体固定宽度
Private Const Frm_LH_gap As Integer = 6 ' 所有控件左对齐的左边距
Private Const itemgap As Integer = 1
' 控件默认尺寸
Private Const cls_H As Integer = 19 ' 高度
Private Const cls_W As Integer = 200 ' 宽度
Private Const Btn_W As Integer = 54 ' 按钮宽度

Private Const Text_W As Integer = 250 ' 输入框宽度（=窗体宽-2*左边距）
Private Const cls_frontsize = 10.5

Private bttop As Integer
Private currTop As Long
' 样式常量（保持美观）
Private Const FONT_NAME As String = "Thoma"
Private Const FONT_SIZE As Integer = 10
Private Const Frm_BACKCOLOR As Long = &H8000000F ' 浅灰背景
Private Const BTN_BACKCOLOR As Long = &H8000000E ' 按钮灰蓝
'
'Private WithEvents ctr As Control
Private lst

Public wdCfg

Option Explicit
Sub setFrm(ttl, inf)
    ' 将变量声明放在开头，更清晰
    Dim cfg As Variant
    Dim thisleft As Long
    Dim ctr As MSForms.Control

    With Me
        .Caption = ttl
        .Width = Frm_WIDTH
        .BackColor = Frm_BACKCOLOR
        .Font.Name = FONT_NAME
        .Font.Size = 10
        .StartUpPosition = 2 ' 居中
        .Height = 100
    End With
    
    bttop = 0
    Set lst = inf
    currTop = 0
       Dim lastH
       
    For Each cfg In lst
        Set ctr = Me.controls.Add(cfg("Type"), cfg("Name"), True)
        With ctr
            .Font.Size = cls_frontsize
            .Name = cfg("Name")
            .Left = Frm_LH_gap
            .Width = cls_W
            Select Case cfg("Type")
                Case "Forms.CommandButton.1"
                    ' -- 修正逻辑开始 --
                    .Width = Btn_W
                    .Height = cls_H ' 使用专用的按钮高度常量
                    If bttop = 0 Then ' 如果是第一个按钮
                        bttop = currTop + 3 * itemgap
                        .top = bttop
                        thisleft = .Left + .Width + itemgap ' 计算下一个按钮的左边距
                        currTop = bttop
                    Else ' 如果是本行的后续按钮
                        .top = bttop
                        .Left = thisleft
                        thisleft = .Left + .Width + itemgap
                        currTop = bttop
                    End If
                Case Else
                    .top = currTop
            End Select
            ' 这部分逻辑保持原样，但为了完整性包含在此
            .Height = cls_H
            If cfg("Type") <> "Forms.TextBox.1" Then
                .Caption = cfg("Caption")
            Else
                .Text = cfg("Caption")
                .Width = Me.Width - 5 * Frm_LH_gap
                .Height = 2 * cls_H
            End If
            currTop = .top + .Height + itemgap
            Debug.Print currTop
          lastH = .Height
        End With
    Next
    
    Me.Height = currTop + lastH + 6 * (itemgap + 1)
End Sub


Private Sub UserForm_Click()

End Sub

