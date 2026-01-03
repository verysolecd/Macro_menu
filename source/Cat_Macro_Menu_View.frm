VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Cat_Macro_Menu_View 
   Caption         =   "UserForm1"
   ClientHeight    =   4035
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   7490
   OleObjectBlob   =   "Cat_Macro_Menu_View.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "Cat_Macro_Menu_View"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' VERSION 5.00
' 窗体边距
Private FrmMargin As Variant ' 上, 右, 下, 左 窗体边距调整值
' 窗体宽度调整值
Private Const ADJUST_F_W = 4
' 窗体高度调整值
Private Const ADJUST_F_H = 5
' 多页控件调整
Private Const ADJUST_M_W = 10 ' 多页控件宽度调整值
Private Const ADJUST_M_H = 5 ' 多页控件高度调整值
Private Const Tab_W = 30 ' Tab固定宽度
Private Const Tab_H = 20 ' TAB高度
Private Const Tab_frontsize = 10
' 按钮尺寸
Private Const Btn_W = 70 ' 按钮的固定宽度
Private Const BTN_H = 20 ' 单个按钮的高度
Private Const BTN_frontsize = 8 ' 按钮字体大小
'标签尺寸
Private Const lb_W = 62 ' 宽度
Private Const lb_H = 26 ' 高度
Private Const lb_frontsize = 10 ' 字体大小
Private Const itl = "公众号:键盘造车手"
' 按钮事件集合
Private mBtns As Object
Private WithEvents prdObserver As Cls_PDM
Attribute prdObserver.VB_VarHelpID = -1
Private WithEvents lblProductInfo As MSForms.Label
Attribute lblProductInfo.VB_VarHelpID = -1
Private WithEvents lblAuthor As MSForms.Label
Attribute lblAuthor.VB_VarHelpID = -1
Private WithEvents MPgs As MSForms.MultiPage
Attribute MPgs.VB_VarHelpID = -1
Option Explicit
' 设置窗体信息
Sub Set_FormInfo(ByVal InfoLst As Object, _
                 ByVal PageMap As Object, _
                 ByVal formTitle As String, _
                 ByVal CloseType As Boolean)
    Set prdObserver = pdm  ' 连接到全局产品观察器
    FrmMargin = Array(2, 2, 2, 2) ' 上, 右, 下, 左 窗体边距调整值
    Set MPgs = Me.controls.Add("Forms.MultiPage.1", "MPgs", True) ' 创建多页控件
    Dim Pgs As Pages
     Set Pgs = MPgs.Pages: Pgs.Clear
    Dim key As Long, KeyStr As Variant, Pg As Page, pName As String
    Dim BtnInfos As Object, info As Variant
    Dim Btns As Object: Set Btns = KCL.InitLst()
    Dim btn As MSForms.CommandButton
    Dim BtnEvt As Cls_btEVT
    For Each KeyStr In InfoLst
        key = CLng(KeyStr)
        If Not PageMap.Exists(key) Then GoTo Continue
        pName = PageMap(key)
        Set Pg = Get_Page(Pgs, pName)
        Set BtnInfos = InfoLst(KeyStr)
        For Each info In BtnInfos
            Set btn = Init_Button(Pg.controls, key, info)
            Set BtnEvt = New Cls_btEVT
            Call BtnEvt.set_ButtonEvent(btn, info, Me, CloseType)
            Btns.Add BtnEvt
        Next
Continue:
    Next
    Set mBtns = Btns
    Call Set_MPage(MPgs)
    Call Set_Form(MPgs, formTitle)
   
    Set lblProductInfo = getNewLbl(Me)
    Set lblAuthor = getMeinfo(Me) ' 创建底部的作者信息栏
    UpdateProductInfo    ' 初始更新产品信息
End Sub
' 设置窗体属性
Private Sub Set_Form(ByVal MPgs As MultiPage, ByVal cap As String)
    With Me
        Dim requiredInsideHeight
        requiredInsideHeight = MPgs.top + MPgs.Height + ADJUST_F_H + lb_H  '+ FrmMargin(2)
        .Height = requiredInsideHeight + (Me.Height - Me.InsideHeight)
        .Width = MPgs.Width + 2 * ADJUST_F_W + 2 * FrmMargin(2)
        .Caption = cap
    End With
End Sub
' 设置多页控件属性
Private Sub Set_MPage(ByVal MPgs As MultiPage)
    MPgs.Width = Tab_W + Btn_W + FrmMargin(3) + ADJUST_M_W
    With MPgs
        .top = lb_H + FrmMargin(1) + 1
        .Left = FrmMargin(0)
        .TabFixedHeight = Tab_H  ' 标签高度（单位：磅）
        .TabFixedWidth = Tab_W ' 标签宽度
        .Font.Name = "Arial"
        .Font.Size = Tab_frontsize
        .MultiRow = True
        .TabOrientation = fmTabOrientationLeft
        .Style = fmTabStyleButtons  ' 切换为按钮样式
     End With
    Dim Pg As Page
    Dim MaxBtnCnt As Long: MaxBtnCnt = 0
    Dim BtnCnt As Long
    For Each Pg In MPgs.Pages
        BtnCnt = Pg.controls.count
        MaxBtnCnt = IIf(BtnCnt > MaxBtnCnt, BtnCnt, MaxBtnCnt)
    Next
    MPgs.Height = FrmMargin(0) + (BTN_H * MaxBtnCnt * 1) + FrmMargin(2) + ADJUST_M_H
    ' 设置多页控件背景颜色
End Sub
' 初始化按钮
Private Function Init_Button(ByVal Ctls As controls, _
                             ByVal idx As Long, _
                             ByVal BtnInfo As Variant) As MSForms.CommandButton
    Dim btn As MSForms.CommandButton
    Set btn = Ctls.Add("Forms.CommandButton.1", idx, True)
    
    Dim Pty As Variant
    For Each Pty In BtnInfo.keys
        Call Try_SetProperty(btn, Pty, BtnInfo.item(Pty))
    Next
    With btn
        .top = (Ctls.count - 1) * BTN_H - 1 '+ (Ctls.Count - 1)+ FrmMargin(0) +
        .Left = FrmMargin(2)
        .Height = BTN_H
        .Width = Btn_W
        .Font.Name = "Arial"   ' 设置按钮字体
        .Font.Size = BTN_frontsize
       ' .BackColor = RGB(220, 220, 220)  ' 设置按钮背景颜色
    End With
    Set Init_Button = btn
End Function
' 尝试设置控件属性
Private Sub Try_SetProperty(ByVal ctrl As Object, _
                            ByVal PptyName As String, _
                            ByVal value As Variant)
    On Error Resume Next
        Err.Number = 0
        Dim tmp As Variant
        tmp = CallByName(ctrl, PptyName, VbGet)
        If Not Err.Number = 0 Then
            Debug.Print PptyName & ": 获取属性失败(" & Err.Number & ")"
            Exit Sub
        End If
        Select Case TypeName(tmp)
            Case "Empty"
                Exit Sub
            Case "Long"
                value = CLng(value)
            Case "Boolean"
                value = CBool(value)
            Case "Currency"
                value = CCur(value)
        End Select
        If Not Err.Number = 0 Then
            Debug.Print value & ": 类型转换失败(" & Err.Number & ")"
            Exit Sub
        End If
        Call CallByName(ctrl, PptyName, VbLet, value)
        If Not Err.Number = 0 Then
            Debug.Print value & ": 设置属性失败(" & Err.Number & ")"
            Exit Sub
        End If
    On Error GoTo 0
End Sub
' 获取页面 - 若不存在则创建
Private Function Get_Page(ByVal Pgs As Pages, ByVal Name As String) As Page
    Dim Pg As Page
    On Error Resume Next
        Set Pg = Pgs.item(Name)
    On Error GoTo 0
    If Pg Is Nothing Then
        Set Pg = Pgs.Add(Name, Name, Pgs.count)
    End If
    Set Get_Page = Pg
End Function
' 产品变化事件处理程序
Private Sub prdObserver_ProductChanged()
 Debug.Print "事件触发"
    UpdateProductInfo
End Sub


' 更新产品信息的方法
Private Sub UpdateProductInfo()
    Dim msg, mcolor
    mcolor = vbRed
    msg = "待选择"
    If Not prdObserver.CurrentProduct Is Nothing Then
          msg = prdObserver.CurrentProduct.PartNumber & "待修改"
          mcolor = vbGreen
    End If
        lblProductInfo.Caption = msg
        lblProductInfo.BackColor = mcolor
End Sub
Private Sub toMP()
    On Error Resume Next
    Dim shell As Object
    Set shell = CreateObject("WScript.Shell")
    shell.Run "https://mp.weixin.qq.com/s?__biz=MzU5MTk1MDUwNg==&mid=2247484525&idx=1&sn=554a37aff4bc876424043a9aa5968d6d&scene=21&poc_token=HCUyg2ijuHYXMx810A5yID4tAYIemJFdJ7FpVvew"
    Set shell = Nothing
    If Err.Number <> 0 Then
        MsgBox "无法公众号链接" & vbCrLf & "错误: " & Err.Description, vbExclamation, "链接错误"
    End If
    On Error GoTo 0
End Sub
Private Sub UserForm_Click()
      toMP
End Sub
Private Sub lblAuthor_Click()
     toMP
End Sub
Private Sub lblProductInfo_Click()
      toMP
End Sub
Private Sub MPgs_MouseDown(ByVal index As Long, ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    If Button <> 1 Then Exit Sub
      If X > Tab_W - 32 Then
      toMP
    End If
End Sub

Private Function getNewLbl(mFrm)
    Dim mLbl
    Set mLbl = mFrm.controls.Add("Forms.Label.1", "lblProductInfo", True)
        With mLbl
             .Caption = "操作产品待选择"
             .top = FrmMargin(0): .Height = lb_H
             .Left = 2: .Width = mFrm.Width - 16
             .Font.Size = lb_frontsize
             .BackColor = vbGreen
             .TextAlign = fmTextAlignCenter
             .BorderStyle = fmBorderStyleSingle
             .WordWrap = True         ' 不换行
             .AutoSize = False
         End With
   Set getNewLbl = mLbl
End Function

Private Function getMeinfo(mFrm)
    Dim mLbl
    Set mLbl = mFrm.controls.Add("Forms.Label.1", "lblAuthor", True)
        With mLbl
            .Caption = itl ' 使用常量显示作者信息
            .top = MPgs.top + MPgs.Height + FrmMargin(1) + 2 ' 放置在多页控件下方
            .Left = lblProductInfo.Left + 5 ' 与顶部信息栏左对齐
            .Width = lblProductInfo.Width ' 与顶部信息栏同宽
            .Height = lb_H
            .Font.Size = lb_frontsize - 1 ' 字体可以稍小一些
            .TextAlign = fmTextAlignCenter
             .WordWrap = False              ' 不换行
             .AutoSize = True
              .BorderStyle = fmBorderStyleSingle
        End With
   Set getMeinfo = mLbl
End Function






