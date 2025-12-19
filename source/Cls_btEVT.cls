VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Cls_btEVT"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private WithEvents mBtn As MSForms.CommandButton
Attribute mBtn.VB_VarHelpID = -1
Private WithEvents btn As MSForms.CommandButton
Attribute btn.VB_VarHelpID = -1

Private mParentForm As MSForms.UserForm
Private mInfo As Object
Private mCloseType As Boolean

Private Const TAG_ENTRYPNT = "ep"           ' 入口点标签
Private Const TAG_PJTPATH = "pjt_path"      ' 项目路径标签
Private Const TAG_MDLNAME = "mdl_name"      ' 模块名称标签


Sub set_ButtonEvent(ByVal btn As MSForms.CommandButton, _
                    ByVal BtnInfo As Object, _
                    ByVal ParentForm As MSForms.UserForm, _
                    ByVal CloseType As Boolean)
    Set mBtn = btn
    Set mInfo = BtnInfo
    Set mParentForm = ParentForm
    mCloseType = CloseType
End Sub
Private Sub mBtn_click()
    If Not mCloseType Then Unload mParentForm  ' 如果是颜色切换按钮，在执行前更新按钮文字
    
    If mInfo(TAG_MDLNAME) = "OTH_Color" Then
        On Error Resume Next
         Dim MyViewer:  Set MyViewer = CATIA.ActiveWindow.ActiveViewer
         Dim currentColor(2)
            MyViewer.GetBackgroundColor currentColor
        On Error GoTo 0
        Dim willBeWhite As Boolean      ' 根据当前背景色判断即将切换到的状态
        willBeWhite = (currentColor(0) <> 1 Or currentColor(1) <> 1 Or currentColor(2) <> 1)
        Call UpdateButtonText(mBtn, willBeWhite)  ' 调用M60_Color模块的函数更新按钮文字
    End If
    
     If mInfo(TAG_MDLNAME) = "OTH_ivhideshow" Then
        On Error Resume Next
        
         ' 根据当前背景色判断即将切换到的状态
        willBeWhite = (currentColor(0) <> 1 Or currentColor(1) <> 1 Or currentColor(2) <> 1)
        Call UpdateButtonText(mBtn, willBeWhite)  ' 调用M60_Color模块的函数更新按钮文字
    End If
    
    Dim ary: ary = Array()  '这里按照传入按钮信息执行按钮对应的module宏
    Dim ss: Set ss = CATIA.SystemService
    Call ss.ExecuteScript(mInfo(TAG_PJTPATH), _
                          catScriptLibraryTypeVBAProject, _
                          mInfo(TAG_MDLNAME), _
                          mInfo(TAG_ENTRYPNT), _
                          Array())
End Sub

Private Sub UpdateButtonText(ByVal btn As MSForms.CommandButton, ByVal isWhiteBackground As Boolean)
    If isWhiteBackground Then
        btn.Caption = "默认背景"
    Else
        btn.Caption = "白色背景"
    End If
End Sub

Private Sub UpdateButtonText2(ByVal btn As MSForms.CommandButton, ByVal isWhiteBackground As Boolean)
    If isWhiteBackground Then
        btn.Caption = "默认背景"
    Else
        btn.Caption = "白色背景"
    End If
End Sub

