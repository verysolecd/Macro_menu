VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} CAT_springWD 
   Caption         =   "UserForm1"
   ClientHeight    =   915
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   1800
   OleObjectBlob   =   "CAT_springWD.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "CAT_springWD"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
    Private Const Frm_LH_gap As Integer = 4 ' 所有控件左对齐的左边距
    Private Const ItmGap As Integer = 0.8
    ' 控件默认尺寸
    Private Const cls_H As Integer = 16 ' 高度
    Private Const cls_W As Integer = 220 ' 宽度
    Private Const Btn_W As Integer = 60 ' 按钮宽度
    Private Const Text_W As Integer = 220
    Private Const cls_frontsize = 11
    Private lst, cfg, ctr
    Private reqHeight, reqWidth, curRH, curBtm, curTop
    Private BtnTop, BtnLeft
    
    ' 样式常量（保持美观）
    Private Const FONT_NAME As String = "Thoma"
    Private Const Frm_color As Long = &H8000000F ' 浅灰背景
    Private Const BTN_color As Long = &H8000000F ' 按钮灰蓝

    Option Explicit
    Private Const mdlName As String = "CAT_springWD"

' --- 主构建函数 ---
Sub setFrm(ttl, cfgs, Optional ByVal isVert As Boolean = False)

    Dim Btnlst As Object: Set Btnlst = KCL.Initlst
    Dim txt_label_lst As Object: Set txt_label_lst = KCL.Initlst
    Dim cfg, ctr As Control
    BtnTop = 0: curTop = 0: BtnLeft = Frm_LH_gap ' 重置全局状态
    
    ' 2. 核心构建循环
    For Each cfg In cfgs
    KCL.showdict cfg
        Set ctr = Me.Controls.Add(cfg("Type"), cfg("Name"), True)
        With ctr
            If Not cfg("Type") = "Forms.TextBox.1" Then .Caption = cfg("Caption")
            .Font.Name = FONT_NAME: .Font.Size = cls_frontsize: .AutoSize = False
            .Height = cls_H: .Width = cls_W:
            On Error Resume Next
                If cfg("Color") <> "" Then .BackColor = KCL.ParseHex(cfg("Color"))
            Error.Clear
            On Error GoTo 0
            ' --- 布局逻辑开始 ---
            If isVert Then  ' === 竖排模式 ===
                .Left = Frm_LH_gap: .Width = Btn_W: .Top = curTop
                Select Case cfg("Type")
                    Case "Forms.TextBox.1"
                        .Height = 2 * cls_H: .AutoSize = False: .text = cfg("Caption")
                        txt_label_lst.Add ctr
                    Case "Forms.Label.1"
                        txt_label_lst.Add ctr
                        .AutoSize = True
                    Case Else
                        .AutoSize = True
                End Select
                curTop = curTop + .Height
            Else  ' === 横排模式 ===
                If cfg("Type") = "Forms.CommandButton.1" Then
                    Btnlst.Add ctr    ' 按钮行暂存
                Else
                    .Left = Frm_LH_gap:  .Top = curTop  ' 非按钮直接放置
                    Select Case cfg("Type")
                        Case "Forms.TextBox.1"
                            .Height = 2 * cls_H: .text = cfg("Caption")
                            txt_label_lst.Add ctr
                        Case "Forms.Label.1", "Forms.Labels.1"
                            .AutoSize = True: txt_label_lst.Add ctr
                        Case Else
                            .AutoSize = True
                    End Select
                    curTop = curTop + .Height + ItmGap
                End If
            End If
            ' --- 布局逻辑结束 ---
        End With
    Next
    
    ' 3. 后处理：横排按钮行
    If Not isVert And Btnlst.count > 0 Then
        curTop = curTop + 3 * ItmGap
        Dim BTN
        For Each BTN In Btnlst
            BTN.Top = curTop: BTN.Height = cls_H
            BTN.Left = BtnLeft: BTN.Width = Btn_W
            BtnLeft = BtnLeft + BTN.Width + 1.5 * ItmGap
        Next
        curTop = curTop + cls_H
    End If
    
    ' 4. 窗体最终定型 (Call Helper)
    FinalizeForm ttl, txt_label_lst
End Sub

' 简单的Helper防止循环内代码太乱
Private Sub ConfigButtons(ctr, col)
    col.Add ctr
End Sub

' --- 独立出来的窗体定型与计算函数 ---
Private Sub FinalizeForm(ttl, txt_label_lst)
    Dim maxW!, maxH!, ctr As Control
    For Each ctr In Me.Controls    ' 计算内容边界
        If ctr.Visible Then
            If ctr.Left + ctr.Width > maxW Then maxW = ctr.Left + ctr.Width
            If ctr.Top + ctr.Height > maxH Then maxH = ctr.Top + ctr.Height
        End If
    Next
    With Me    ' 设置窗体属性
        .Caption = ttl: .BackColor = Frm_color
        .Font.Name = FONT_NAME: .Font.Size = cls_frontsize
        .StartUpPosition = 2
        .Width = maxW + (.Width - .InsideWidth) + Frm_LH_gap
        .Height = maxH + (.Height - .InsideHeight) + 6 * ItmGap
    End With
    Dim t
    For Each t In txt_label_lst ' 自适应拉伸文本框
        t.Width = Me.InsideWidth - 2 * Frm_LH_gap
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
      toMP
End Sub
Private Sub lbL_jpzcs_Click()
     toMP
End Sub

