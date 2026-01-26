VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cls_dynaFrmMLess"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

' ==============================================================================
' cls_dynaFrmMLess
' ------------------------------------------------------------------------------
' 功能：通用的非模态窗体管理器
' 作用：动态解析配置，使用 springWD 渲染非模态工具条，并管理事件绑定。
' 作者：Antigravity Agent
' ==============================================================================
Option Explicit
Private mForm As springWD               ' 动态窗体实例
Private mEventCol As collection         ' 事件处理器集合（防止对象销毁）
Private mConfigLst As Object            ' 控件配置列表 (KCL.InitLst)

' --- 初始化 ---
Private Const mdlname As String = "cls_dynaFrmMLess"
Private Sub Class_Initialize()
    Set mEventCol = New collection
End Sub

Private Sub Class_Terminate()
    ' 当类被销毁时，即表示管理器退出，此时应关闭窗体
    On Error Resume Next
    If Not mForm Is Nothing Then
        Unload mForm
        Set mForm = Nothing
    End If
    Set mEventCol = Nothing
    On Error GoTo 0
End Sub

' ==============================================================================
' ShowToolbar
' ------------------------------------------------------------------------------
' 参数：
'   Title     : 窗体标题
'   ConfigStr : 类似 cls_dynafrm 的配置字符串 (包含 %UI 定义)
'   Project   : 宏所在的项目名 (如 CATIA_V5_SimpleMacroMenu.catvba)
'   ModuleMap : 字典，Key=按钮名, Value=对应的模块名(如 RW_4readPrd)
'   MacroMap  : 字典，Key=按钮名, Value=对应的宏名(如 readPrd)
' ==============================================================================
Public Sub ShowToolbar(ByVal Title As String, _
                       ByVal ConfigStr As String, _
                       ByVal Project As String, _
                       ByVal ModuleMap As Object, _
                       ByVal MacroMap As Object)
                       
    ' 1. 解析配置字符串
    Set mConfigLst = ParseConfig(ConfigStr)
    ' 2. 创建并渲染窗体
    If Not mForm Is Nothing Then Unload mForm
    Set mForm = New springWD
    mForm.setFrm Title, mConfigLst
    ' 3. 绑定事件 (核心逻辑)
    BindEvents mForm, Project, ModuleMap, MacroMap
    ' 4. 显示非模态
    mForm.Show vbModeless
    
End Sub

' ------------------------------------------------------------------------------
' 辅助：解析配置 (复用 cls_dynafrm 的逻辑简化版)
' ------------------------------------------------------------------------------
Private Function ParseConfig(ByVal code As String) As Object
    Dim regEx As Object, matches As Object, match As Object
    Dim item As Object, lst As Object
    
    Set lst = KCL.InitLst
    Set regEx = CreateObject("VBScript.RegExp")
    
    ' 正则匹配: ' %UI <Type> <Name> <Caption>
    With regEx
        .Global = True
        .MultiLine = True
        .Pattern = "^\s*'\s*%UI\s+(\w+)\s+(\w+)\s+(.*)$"
    End With
    
    If regEx.TEST(code) Then
        Set matches = regEx.Execute(code)
        For Each match In matches
            Set item = KCL.InitDic
            item.Add "Type", GetControlType(match.SubMatches(0))
            item.Add "Name", match.SubMatches(1)
            item.Add "Caption", VBA.Trim(match.SubMatches(2))
            lst.Add item
        Next
    End If
    Set ParseConfig = lst
End Function

' ------------------------------------------------------------------------------
' 辅助：绑定事件
' ------------------------------------------------------------------------------
Private Sub BindEvents(frm As Object, pjt As String, modMap As Object, macMap As Object)
    Dim ctrl As Control
    Dim btnName As String
    Dim handler As Cls_btEVT
    Dim info As Object
    
    Set mEventCol = New collection
    
    For Each ctrl In frm.Controls
        If TypeName(ctrl) = "CommandButton" Then
            btnName = ctrl.Name
            
            ' 如果配置了该按钮的映射
            If modMap.Exists(btnName) And macMap.Exists(btnName) Then
                
                ' 构造 Cls_btEVT 需要的 info 字典
                Set info = CreateObject("Scripting.Dictionary")
                info.Add "pjt_path", pjt
                info.Add "mdl_name", modMap(btnName)
                info.Add "ep", macMap(btnName)
                
                ' 创建事件处理器并绑定
                Set handler = New Cls_btEVT
                ' set_ButtonEvent: btn, info, parent, closeAfterClick
                handler.set_ButtonEvent ctrl, info, frm, True ' False表示点击后不关闭窗体
                
                ' 加入集合保活
                mEventCol.Add handler
            End If
        End If
    Next
End Sub

' ------------------------------------------------------------------------------
' 辅助：控件类型映射
' ------------------------------------------------------------------------------
Private Function GetControlType(ByVal alias As String) As String
    alias = LCase(alias)
    Select Case alias
        Case "button", "btn", "cmd": GetControlType = "Forms.CommandButton.1"
        Case "label", "lbl": GetControlType = "Forms.Label.1"
        Case "text", "txt", "textbox": GetControlType = "Forms.TextBox.1"
        Case Else: GetControlType = "Forms.Label.1"
    End Select
End Function




