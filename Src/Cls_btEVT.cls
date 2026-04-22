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
' --- Variables ---
Private WithEvents mBTN As MSForms.CommandButton      ' For Script Mode
Attribute mBTN.VB_VarHelpID = -1
Private WithEvents mBTNwd As MSForms.CommandButton ' For Result Mode
Attribute mBTNwd.VB_VarHelpID = -1


Private mInfo As Object
Private mCloseType As Boolean
Private mParentForm As Object   ' Used by Script Mode
Private mParentFrm As Object    ' Used by Result Mode (Unified target)

Private mControlID As String
Private mResultDict As Object

Private Const TAG_ENTRYPNT = "ep"           ' 入口点标签
Private Const TAG_PJTPATH = "pjt_path"      ' 项目路径标签
Private Const TAG_MDLNAME = "mdl_name"      ' 模块名称标签
Private Const TAG_PASS_BTN = "springwdBTN"  ' 传递按钮名称开关
Private Const mdlname As String = "Cls_btEVT"

' For MLess / Script Mode
Sub set_ButtonEvent(ByVal BTN As MSForms.CommandButton, _
                    ByVal BtnInfo As Object, _
                    ByVal ParentForm As Object, _
                    ByVal CloseType As Boolean)
        Set mBTN = BTN
        Set mInfo = BtnInfo
        Set mParentForm = ParentForm
        mCloseType = CloseType
End Sub

' For Result Mode / Frm
Public Property Set ParentFrm(ByVal frm As Object)
    Set mParentFrm = frm
End Property
Public Property Let ControlID(ByVal Value As String)
    mControlID = Value
End Property
Public Property Set ControlBTN(ByVal BTN As MSForms.CommandButton)
    Set mBTNwd = BTN
End Property
Public Property Set ResultDict(ByVal dict As Object)
    Set mResultDict = dict
End Property
' Handler for Result Mode (mBTNwd) 模态按钮的处理
Private Sub mBTNwd_Click()
    If mParentFrm Is Nothing Then Exit Sub
        mParentFrm.Tag = mControlID
    If Not mResultDict Is Nothing Then
        Dim ctl As Control
        For Each ctl In mParentFrm.Controls
            On Error Resume Next
            If TypeName(ctl) = "CheckBox" Then
                 mResultDict(ctl.Name) = ctl.Value
            ElseIf TypeName(ctl) = "TextBox" Then
                 mResultDict(ctl.Name) = ctl.text
            End If
            On Error GoTo 0
        Next
    End If
    mParentFrm.Hide
End Sub
' Handler for Script Mode (mBTN) 非模态按钮的处理
Private Sub mBTN_click()
    If Not mCloseType Then
        If Not mParentForm Is Nothing Then Unload mParentForm
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
        params = Array(mBTN.Name)
    Else
        params = Array()
    End If
    
    On Error Resume Next
        Set g_Btn = mBTN
        Call ss.ExecuteScript(mInfo(TAG_PJTPATH), _
                              catScriptLibraryTypeVBAProject, _
                              mInfo(TAG_MDLNAME), _
                              mInfo(TAG_ENTRYPNT), _
                              params)
         Err.Clear
    On Error GoTo 0
End Sub








