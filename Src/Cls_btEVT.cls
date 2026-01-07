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

    If Not mCloseType Then Unload mParentForm
    Dim ary: ary = Array()  '这里按照传入按钮信息执行按钮对应的module宏
    Dim ss: Set ss = CATIA.SystemService
    Call ss.ExecuteScript(mInfo(TAG_PJTPATH), _
                          catScriptLibraryTypeVBAProject, _
                          mInfo(TAG_MDLNAME), _
                          mInfo(TAG_ENTRYPNT), _
                          Array())
                          
    Dim capUp, icap
    icap = Array(0, 0)
    Select Case VBA.UCase(mInfo(TAG_MDLNAME))
    Case "OTH_COLOR"
            capUp = amWhite
            icap = Array("白色背景", "默认背景")
    Case "OTH_SWSCR"
             capUp = canrefresh
           
            icap = Array("切换:禁止更新", "切换:正常模式")
    End Select
    If icap(0) <> 0 Then Call UpdateButtonText(capUp, icap)
               
End Sub

Private Sub UpdateButtonText(ByVal up2New As Boolean, ary)
        mBtn.Caption = IIf(up2New, ary(0), ary(1))
        MsgBox "CATIA.RefreshDisplay=" & CATIA.RefreshDisplay & vbNewLine _
                                & "点击后" & mBtn.Caption

End Sub
Private Function amWhite()
          amWhite = False
        If CATIA.Windows.count >= 1 Then
            Dim currentColor(2)
                Dim MyViewer:  Set MyViewer = CATIA.ActiveWindow.ActiveViewer
                MyViewer.GetBackgroundColor currentColor
            amWhite = (currentColor(0) <> 1 Or currentColor(1) <> 1 Or currentColor(2) <> 1) '只要一个不等于1就是true
        End If
End Function

Private Function canrefresh()

Dim bol: bol = CATIA.RefreshDisplay
If bol Then
KCL.setASM (bol)
Else
KCL.setASM (bol)
End If

CATIA.ActiveWindow.ActiveViewer.Update
canrefresh = CATIA.RefreshDisplay
End Function






