VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Button_Evt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'VERSION 1.0 CLASS
'Begin
'  MultiUse = -1  'True
'End
'Attribute VB_Name = "Button_Evt"
'Attribute VB_GlobalNameSpace = False
'Attribute VB_Creatable = False
'Attribute VB_PredeclaredId = False
'Attribute VB_Exposed = False
'' Button_Evt.cls
'' 此代码用于处理按钮事件

'Attribute mBtn.VB_VarHelpID = -1
' Cat_Macro_Menu_Model.bas
' 定义全局变量（防止命名冲突）

Option Explicit


Private WithEvents mBtn As MSForms.CommandButton
Attribute mBtn.VB_VarHelpID = -1
Private mInfo As Object
Private mParentForm As MSForms.UserForm
Private mCloseType As Boolean
Private Const TAG_ENTRYPNT = "ep"           ' 入口点标签
Private Const TAG_PJTPATH = "pjt_path"      ' 项目路径标签
Private Const TAG_MDLNAME = "mdl_name"      ' 模块名称标签

' 设置事件
Sub set_ButtonEvent(ByVal btn As MSForms.CommandButton, _
              ByVal BtnInfo As Object, _
              ByVal ParentForm As MSForms.UserForm, _
              ByVal CloseType As Boolean)
    Set mBtn = btn
    Set mInfo = BtnInfo
    Set mParentForm = ParentForm
    mCloseType = CloseType
End Sub
' 按钮点击事件处理
'Private Sub mBtn_Click()
'    If Not mCloseType Then Unload mParentForm
'
'    Dim ary: ary = Array()
'    Dim ss: Set ss = CATIA.SystemService
'
'    Call ss.ExecuteScript(mInfo(TAG_PJTPATH), _
'                          catScriptLibraryTypeVBAProject, _
'                          mInfo(TAG_MDLNAME), _
'                          mInfo(TAG_ENTRYPNT), _
'                          Array())
'End Sub


Private Sub mBtn_Click()
    If Not mCloseType Then Unload mParentForm
    
    ' 如果是颜色切换按钮，在执行前更新按钮文字
    If mInfo(TAG_MDLNAME) = "M60_Color" Then
        Dim MyViewer As Object
        Set MyViewer = CATIA.activeWindow.ActiveViewer
        Dim currentColor(2)
        MyViewer.GetBackgroundColor currentColor
        
        ' 根据当前背景色判断即将切换到的状态
        Dim willBeWhite As Boolean
        willBeWhite = (currentColor(0) <> 1 Or currentColor(1) <> 1 Or currentColor(2) <> 1)
        
        ' 调用M60_Color模块的函数更新按钮文字
        Call M60_Color.UpdateButtonText(mBtn, willBeWhite)
    End If
    
    Dim ary: ary = Array()
    Dim ss: Set ss = CATIA.SystemService
    
    Call ss.ExecuteScript(mInfo(TAG_PJTPATH), _
                          catScriptLibraryTypeVBAProject, _
                          mInfo(TAG_MDLNAME), _
                          mInfo(TAG_ENTRYPNT), _
                          Array())
End Sub
