VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Cls_allBTNEVT"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
' --- Variables ---
Private WithEvents mBTN_SM As MSForms.CommandButton      ' For Script Mode
Attribute mBTN_SM.VB_VarHelpID = -1
Private WithEvents mBTN_RM As MSForms.CommandButton ' For Result Mode
Attribute mBTN_RM.VB_VarHelpID = -1

Private mFrm_SM As Object   ' [Script Mode] 工具栏窗体引用（用于 Unload） Private mFrm_Script As Object
Private mFrm_RM As Object   ' [Result Mode] 模态窗体引用（用于 Hide + 收集结果） Private mFrm_Result As Object

Private mInfo As Object
Private mKeepOpen As Boolean

Private mControlID As String
Private mResultDict As Object

Private Const TAG_ENTRYPNT = "ep"           ' 入口点标签
Private Const TAG_PJTPATH = "pjt_path"      ' 项目路径标签
Private Const TAG_MDLNAME = "mdl_name"      ' 模块名称标签
Private Const TAG_PASS_BTN = "springwdBTN"  ' 传递按钮名称开关

Private Const mdlname As String = "Cls_allBTNEVT"
Public Property Set ParentFrm(ByVal frm As Object)
    Set mFrm_RM = frm
End Property
Public Property Let ControlID(ByVal Value As String)
    mControlID = Value
End Property
Public Property Set ControlBTN(ByVal btn As MSForms.CommandButton)
    Set mBTN_RM = btn
End Property
Public Property Set ResultDict(ByVal dict As Object)
    Set mResultDict = dict
End Property
' For MLess / Script Mode
Sub set_ButtonEvent(ByVal btn As MSForms.CommandButton, _
                    ByVal BtnInfo As Object, _
                    ByVal ParentForm As Object, _
                    ByVal keepOpen As Boolean)
        Set mBTN_SM = btn
        Set mInfo = BtnInfo
        Set mFrm_SM = ParentForm
        mKeepOpen = keepOpen
End Sub
' For Result Mode / Frm

'模态弹窗（结果模式）' Handler for Result Mode (mBTN_RM) 模态按钮的处理
Private Sub mBTN_RM_Click()
    If mFrm_RM Is Nothing Then Exit Sub
        mFrm_RM.Tag = mControlID
    If Not mResultDict Is Nothing Then
        Dim ctl As Control
        For Each ctl In mFrm_RM.Controls
            On Error Resume Next
            If TypeName(ctl) = "CheckBox" Then
                 mResultDict(ctl.Name) = ctl.Value
            ElseIf TypeName(ctl) = "TextBox" Then
                 mResultDict(ctl.Name) = ctl.text
            End If
            On Error GoTo 0
        Next
    End If
    mFrm_RM.Hide
End Sub

' Handler for Script Mode (mBTN_SM) 非模态按钮的处理
Private Sub mBTN_SM_click()
    If Not mKeepOpen Then
        If Not mFrm_SM Is Nothing Then Unload mFrm_SM
    End If
'Func ExecuteScript( iLibraryName,iLibraryType,iProgramName,iFunctionName,iParameters)
'    iLibraryName: The library in which the script is contained
'    iLibraryType: The type of the library
'    iProgramName: The name of the program in the library
'    iFunctionName:The name of the function to invoke
'    IParameters: An array of parameters for the function
    Dim ss As Object: Set ss = CATIA.SystemService
    Dim params As Variant
    If mInfo.Exists(TAG_PASS_BTN) And mInfo(TAG_PASS_BTN) = True Then
        params = Array(mBTN_SM.Name)
    Else
        params = Array()
    End If
    On Error Resume Next
        Set g_Btn = mBTN_SM
        Call ss.ExecuteScript(mInfo(TAG_PJTPATH), _
                              catScriptLibraryTypeVBAProject, _
                              mInfo(TAG_MDLNAME), _
                              mInfo(TAG_ENTRYPNT), _
                              params)
         Err.Clear
    On Error GoTo 0
End Sub
