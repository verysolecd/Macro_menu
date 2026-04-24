VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Cls_DynaUIEngine"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

'===========================================================================
' Cls_DynaUIEngine
' 统一的动态UI引擎 —— 集按钮生成、事件绑定、结果收集于一体
'
' 三种数据加载模式:
'   1. LoadFromModuleName  —— 解析模块头部 %UI / %Title 注释 (二级弹窗)
'   2. LoadFromMenuTags    —— 扫描整个VBA工程的 {GP:}{EP:} 标签 (主菜单)
'   3. AddUIElement        —— 纯代码手动追加控件定义 (扩展)
'
' 两种显示模式:
'   Show       —— 模态/非模态弹窗 (替代 cls_dynaFrm.Show)
'   ShowToolbar—— 非模态工具栏     (替代 cls_dynaFrm.ShowToolbar)
'   Alert      —— 信息弹窗         (替代 cls_dynaFrm.Alert)
'
' 结果回收:
'   Results / ClickedButton / IsCancelled

'类的核心结构
'3 种 UI 生成方式（核心）
'按钮回调完整流程 (核心)
'窗体显示与结果回收
'关键内部机制

'变量 作用

'mWD 动态窗体实例(CAT_springWD)
'mEventCol   事件处理器集合（防止被 GC 回收）
'mResultDict 模态窗体结果字典(点击按钮?输入值)
'mIsAlertMode 提示框模式开关

'===========================================================================
Option Explicit
' --- 控件定义列表 (ArrayList of Dictionary) ---
Private mCtrlDefs As Object       ' 每个元素: Dict{Type, Name, Caption, Color}
' --- 窗体标题 ---
Private mTitle As String
Private mWD As Object             ' CAT_springWD 窗体实例
Private mEventCol As Collection
Private mResultDict As Object
Private mExecuted As Boolean
Private mIsVertical As Variant    ' Empty=Auto, True/False=Manual
Private mPassBtn As Boolean
Private mIsAlertMode As Boolean
Private mAlertMsg As String
Private mPjtPath As String

' --- 剪贴板相关API (Alert复制功能) ---
Private Declare PtrSafe Function OpenClipboard Lib "user32" (ByVal hwnd As LongPtr) As Long
Private Declare PtrSafe Function EmptyClipboard Lib "user32" () As Long
Private Declare PtrSafe Function CloseClipboard Lib "user32" () As Long
Private Declare PtrSafe Function SetClipboardData Lib "user32" (ByVal wFormat As Long, ByVal hMem As LongPtr) As LongPtr
Private Declare PtrSafe Function GlobalAlloc Lib "kernel32" (ByVal wFlags As Long, ByVal dwBytes As LongPtr) As LongPtr
Private Declare PtrSafe Function GlobalLock Lib "kernel32" (ByVal hMem As LongPtr) As LongPtr
Private Declare PtrSafe Function GlobalUnlock Lib "kernel32" (ByVal hMem As LongPtr) As Long
Private Declare PtrSafe Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (ByVal Destination As LongPtr, ByVal Source As Any, ByVal Length As LongPtr)

Private Const GMEM_MOVEABLE = &H2
Private Const CF_TEXT = 1

' --- Alert模式的复制按钮 (需要 WithEvents) ---
Private WithEvents mCopyBtn As MSForms.CommandButton
Attribute mCopyBtn.VB_VarHelpID = -1

' --- 标签常量 ---
Private Const TAG_ENTRYPNT = "ep"
Private Const TAG_PJTPATH = "pjt_path"
Private Const TAG_MDLNAME = "mdl_name"
Private Const TAG_PASS_BTN = "springwdBTN"

' ═══════════════════════════════════════════════════════════════
'  Lifecycle
' ═══════════════════════════════════════════════════════════════

Private Sub Class_Initialize()
    Set mCtrlDefs = P_Initlst()
    Set mEventCol = New Collection
    Set mResultDict = P_InitDic()
    mTitle = "请问如何执行"
    mExecuted = False
    mIsVertical = Empty
    mPassBtn = False
    mIsAlertMode = False
    mAlertMsg = ""
    mPjtPath = ""
    On Error Resume Next
        mPjtPath = P_GetVBAprj()
    On Error GoTo 0
End Sub

Private Sub Class_Terminate()
    On Error Resume Next
    If Not mWD Is Nothing Then Unload mWD
    Set mWD = Nothing
    Set mCtrlDefs = Nothing
    Set mEventCol = Nothing
    Set mResultDict = Nothing
    Set mCopyBtn = Nothing
    ' 注意: 不清除 g_frm，那是老的 cls_dynaFrm 系统的全局变量
    mExecuted = False
    mIsAlertMode = False
    Err.Clear
    On Error GoTo 0
End Sub

' ═══════════════════════════════════════════════════════════════
'  Section A: 配置属性 (Configuration Properties)
''设置/获取窗体标题
Public Property Let Title(ByVal val As String)
    mTitle = val
End Property
Public Property Get Title() As String
    Title = mTitle
End Property
''设置按钮排列方向：True=竖排, False=横排, 不设置=自动
Public Property Let isVertical(ByVal val As Boolean)
    mIsVertical = val
End Property
''设置是否向被调用宏传递按钮名称
Public Property Let PassButtonName(ByVal val As Boolean)
    mPassBtn = val
End Property
'  Section B: 数据加载 —— 三种入口
'' 入口1: 从模块名称加载 (解析 %UI / %Title 注释)
Public Sub LoadFromModuleName(ByVal modName As String)
    Dim rawCode As String
    rawCode = P_getbf1stproc(modName)
    If rawCode = "" Then Exit Sub
    mTitle = ExtractTitle(rawCode)
    Set mCtrlDefs = ExtractUIControls(rawCode)
End Sub



''────────────────────────────────────────────────────────────────
'' 入口3: 纯代码手动追加一个控件定义
''        适合在运行时动态扩展UI，无需任何注释标签
''
'' @param CtrlType  控件类型别名: Button/Label/TextBox/CheckBox
'' @param CtrlName  控件唯一名称 (事件中使用此名做回调标识)
'' @param Caption   显示文字
'' @param HexColor  可选16进制色值 如 "#FF6600"
''────────────────────────────────────────────────────────────────
Public Sub AddUIElement(ByVal CtrlType As String, _
                        ByVal CtrlName As String, _
                        ByVal Caption As String, _
                        Optional ByVal HexColor As String = "")
    Dim def As Object: Set def = P_InitDic()
    def.Add "Type", ResolveCtrlType(CtrlType)
    def.Add "Name", CtrlName
    def.Add "Caption", Caption
    def.Add "Color", HexColor
    mCtrlDefs.Add def
End Sub
' ═══════════════════════════════════════════════════════════════
'  Section C: 显示 —— Show / ShowToolbar / Alert
'' 通用显示入口 (模态 / 非模态)
''
'' @param iShowMode  vbModal 或 vbModeless
'' @param ModuleMap  非模态模式: Dict{btnName -> moduleName}
'' @param MacroMap   非模态模式: Dict{btnName -> entryPoint}
''────────────────────────────────────────────────────────────────
Public Sub Show(Optional ByVal iShowMode As FormShowConstants = vbModal, _
                Optional ByVal ModuleMap As Object = Nothing, _
                Optional ByVal MacroMap As Object = Nothing)
    If mExecuted And iShowMode = vbModal Then Exit Sub
    If iShowMode = vbModal Then mExecuted = True
    If mCtrlDefs.count < 1 Then Exit Sub
    
    ' 创建窗体
    If Not mWD Is Nothing Then Unload mWD
    Set mWD = P_CreateSpringWD()
    
    ' 决定排列方向
    Dim btnVert As Boolean
    If IsEmpty(mIsVertical) Then
        btnVert = (iShowMode = vbModeless)
    Else
        btnVert = mIsVertical
    End If
    
    ' 渲染控件到窗体
    mWD.setFrm mTitle, mCtrlDefs, btnVert
    
    If mIsAlertMode Then
        SetupAlertMode
    Else
        BindEvents mWD, ModuleMap, MacroMap
    End If
    
    ' 显示
    mWD.Show iShowMode
    
    ' 模态窗关闭后收集结果
    If iShowMode = vbModal And mIsAlertMode = False Then
        If mWD.Tag = "" Then
            mWD.Tag = "UserClosed": Unload mWD: Set mWD = Nothing
            Exit Sub
        End If
        If mResultDict Is Nothing Then Set mResultDict = P_InitDic()
        mResultDict.Add "btn_clicked", mWD.Tag
        mWD.Hide: Unload mWD: Set mWD = Nothing
    End If
End Sub

''便捷方法: 显示为非模态工具栏 (等效于原 cls_dynaFrm.ShowToolbar)
'' modMap/macMap 可选:
''   不传 → 自动按约定构建: btnName→modName, btnName→"btnName_click"
''   传入 → 使用自定义映射 (如全部指向 Action_ClickHandler)
Public Sub ShowToolbar(ByVal modName As String, _
                       Optional ByVal modMap As Object = Nothing, _
                       Optional ByVal macMap As Object = Nothing)
    LoadFromModuleName modName
    Me.isVertical = True
    ' 未传映射时自动构建
    If modMap Is Nothing Or macMap Is Nothing Then
        Dim autoMdl As Object, autoFunc As Object
        BuildDefaultMaps modName, autoMdl, autoFunc
        If modMap Is Nothing Then Set modMap = autoMdl
        If macMap Is Nothing Then Set macMap = autoFunc
    End If
    Show vbModeless, modMap, macMap
End Sub

''便捷方法: 显示信息弹窗 (等效于原 cls_dynaFrm.Alert)
Public Sub Alert(ByVal Content As String)
    mIsAlertMode = True
    mAlertMsg = Content
    mExecuted = False
    Show vbModal
End Sub
' ═══════════════════════════════════════════════════════════════
'  Section D: 结果回收 (Result Properties)
''获取所有控件的值 (文本框/复选框等) + btn_clicked
Public Property Get Results() As Object
    EnsureExecuted
    Set Results = mResultDict
End Property
''获取被点击的按钮名称
Public Property Get ClickedButton() As Variant
    EnsureExecuted
    ClickedButton = IIf(mResultDict.Exists("btn_clicked"), mResultDict("btn_clicked"), "")
End Property
''用户是否取消/关闭了窗体
Public Property Get IsCancelled() As Boolean
    EnsureExecuted
    IsCancelled = IIf(mResultDict.Exists("btn_clicked") And mResultDict("btn_clicked") <> "UserClosed", False, True)
End Property
' ═══════════════════════════════════════════════════════════════
'  Internal: 解耦后的 UI 解析子函数 (从 KCL.ParseUIConfig 拆出)
'' 从原始代码中提取标题 (%Title 行)
'' @param  rawCode  模块声明区原始文本
'' @return          标题字符串, 无则返回默认值
''────────────────────────────────────────────────────────────────
Private Function ExtractTitle(ByVal rawCode As String) As String
    ExtractTitle = "请问如何执行"    ' 默认值
    Dim regEx As Object: Set regEx = P_getRegexp()
    With regEx
        .Global = True
        .MultiLine = True
        .Pattern = "^\s*'\s*%Title\s+(.*)$"
    End With
    If Not regEx.Test(rawCode) Then Exit Function
    
    Dim matches As Object: Set matches = regEx.Execute(rawCode)
    Dim match As Object
    For Each match In matches
        ExtractTitle = match.SubMatches(0)
    Next
End Function

''────────────────────────────────────────────────────────────────
'' 从原始代码中提取所有 %UI 控件定义
'' @param  rawCode  模块声明区原始文本
'' @return          ArrayList, 每个元素为 Dict{Type, Name, Caption, Color}
''────────────────────────────────────────────────────────────────
Private Function ExtractUIControls(ByVal rawCode As String) As Object
    Set ExtractUIControls = P_Initlst()
    
    Dim lines As Variant
    lines = Split(rawCode, vbLf)
    
    Dim i As Long
    Dim def As Object
    For i = LBound(lines) To UBound(lines)
        Set def = ParseSingleUILine(CStr(lines(i)))
        If Not def Is Nothing Then
            ExtractUIControls.Add def
        End If
    Next
End Function
''────────────────────────────────────────────────────────────────
'' 解析单行 %UI 注释
'' 格式:  ' %UI <ControlType> <ControlName> <Caption> [#HexColor]
'' 例如:  ' %UI Button btnOK  确定 #FF6600
'' @param  sLine  单行文本
'' @return        Dict{Type, Name, Caption, Color} 或 Nothing
''────────────────────────────────────────────────────────────────
Private Function ParseSingleUILine(ByVal sLine As String) As Object
    Set ParseSingleUILine = Nothing
    sLine = Trim(Replace(sLine, vbCr, "")) ' 去掉行首尾空白和回车
    If Left(sLine, 1) <> "'" Then Exit Function  ' 快速排除: 必须以 ' 开头且包含 %UI
    If InStr(1, sLine, "%UI", vbTextCompare) < 1 Then Exit Function
    Dim regEx As Object: Set regEx = P_getRegexp() ' 用正则精确匹配
    With regEx
        .Global = False
        .MultiLine = False
        .Pattern = "^\s*'\s*%UI\s+(\w+)\s+(\w+)\s+(.+?)(?:\s+(#[0-9a-fA-F]{6}))?\s*$"
    End With
    If Not regEx.Test(sLine) Then Exit Function
    Dim matches As Object: Set matches = regEx.Execute(sLine)
    Dim m As Object: Set m = matches(0)
    Set ParseSingleUILine = P_InitDic()
    With ParseSingleUILine
        .Add "Type", ResolveCtrlType(CStr(m.SubMatches(0)))
        .Add "Name", CStr(m.SubMatches(1))
        .Add "Caption", Trim(CStr(m.SubMatches(2)))
        .Add "Color", Trim(CStr(m.SubMatches(3)))
    End With
End Function
''────────────────────────────────────────────────────────────────
'' 将控件类型别名转换为 MSForms ProgID''
'' @param alias  Button/btn/cmd/Label/lbl/Text/txt/CheckBox/chk
'' @return       "Forms.CommandButton.1" 等
''────────────────────────────────────────────────────────────────
Private Function ResolveCtrlType(ByVal alias As String) As String
    Select Case LCase(alias)
        Case "button", "btn", "cmd":    ResolveCtrlType = "Forms.CommandButton.1"
        Case "label", "lbl":            ResolveCtrlType = "Forms.Label.1"
        Case "text", "txt", "textbox":  ResolveCtrlType = "Forms.TextBox.1"
        Case "checkbox", "chk":         ResolveCtrlType = "Forms.CheckBox.1"
        Case Else:                      ResolveCtrlType = "Forms.Label.1"
    End Select
End Function

' ═══════════════════════════════════════════════════════════════
'  Internal: 菜单扫描辅助函数


' ═══════════════════════════════════════════════════════════════
'  Internal: 事件绑定
' ═══════════════════════════════════════════════════════════════

''为窗体中的所有按钮绑定事件处理
Private Sub BindEvents(frm As Object, modMap As Object, macMap As Object)
    Dim ctl As Control, btnName As String
    Dim handler As Cls_MnuBtnEVT, mBTNinfo As Object
    Dim isScript As Boolean
    isScript = (Not modMap Is Nothing) And (Not macMap Is Nothing)
    
    Set mEventCol = New Collection
    For Each ctl In frm.Controls
        If TypeName(ctl) = "CommandButton" Then
            btnName = ctl.Name
            Set handler = New Cls_MnuBtnEVT
            If isScript And Not modMap Is Nothing Then
                If modMap.Exists(btnName) And macMap.Exists(btnName) Then
                    Set mBTNinfo = P_InitDic()
                    mBTNinfo.Add TAG_PJTPATH, mPjtPath
                    mBTNinfo.Add TAG_MDLNAME, modMap(btnName)
                    mBTNinfo.Add TAG_ENTRYPNT, macMap(btnName)
                    If mPassBtn Then mBTNinfo.Add TAG_PASS_BTN, True
                    handler.set_ButtonEvent ctl, mBTNinfo, frm, True
                End If
            Else
                Set handler.ControlBTN = ctl
                handler.ControlID = btnName
                Set handler.ParentFrm = frm
                Set handler.ResultDict = mResultDict
            End If
            mEventCol.Add handler
        End If
    Next
End Sub

' ═══════════════════════════════════════════════════════════════
'  Internal: Alert 模式设置
' ═══════════════════════════════════════════════════════════════

''将窗体切换为信息展示模式
Private Sub SetupAlertMode()
    Dim ctl As Control
    mWD.Tag = ""
    For Each ctl In mWD.Controls
        If TypeName(ctl) = "CommandButton" Then
            ctl.Visible = False
        ElseIf TypeName(ctl) = "TextBox" Then
            ctl.Value = mAlertMsg
        End If
    Next
    
    ' 动态添加"复制并关闭"按钮
    Dim btnCopy As MSForms.CommandButton
    Set btnCopy = mWD.Controls.Add("Forms.CommandButton.1", "btn_Copy", True)
    With btnCopy
        .Caption = "复制并关闭"
        .Width = 80
        .Height = 24
        .Left = (mWD.InsideWidth - .Width) / 2
        .Top = mWD.InsideHeight - .Height - 6
    End With
    Set mCopyBtn = btnCopy
End Sub

''Alert模式复制按钮点击事件
Private Sub mCopyBtn_Click()
    On Error Resume Next
    Dim hGlobal As LongPtr, lpGlobal As LongPtr
    Dim sText As String: sText = mAlertMsg & vbNullChar
    
    hGlobal = GlobalAlloc(GMEM_MOVEABLE, LenB(sText))
    If hGlobal <> 0 Then
        lpGlobal = GlobalLock(hGlobal)
        If lpGlobal <> 0 Then
            CopyMemory lpGlobal, StrPtr(StrConv(sText, vbFromUnicode)), LenB(StrConv(sText, vbFromUnicode))
            GlobalUnlock hGlobal
            If OpenClipboard(0) <> 0 Then
                EmptyClipboard
                SetClipboardData CF_TEXT, hGlobal
                CloseClipboard
            End If
        End If
    End If
    Unload mWD
End Sub

' ═══════════════════════════════════════════════════════════════
'  Internal: 辅助
' ═══════════════════════════════════════════════════════════════

''确保模态窗已执行过 (惰性执行)
Private Sub EnsureExecuted()
    If Not mExecuted And mCtrlDefs.count > 0 Then Show vbModal
End Sub

''────────────────────────────────────────────────────────────────
'' 从已解析的 mCtrlDefs 自动构建按钮映射
'' 等效于原 KCL.setBTNmdl + KCL.setBTNFunc，但无需重新解析
''
'' 默认约定:
''   modMap: {btnName → modName}       (所有按钮属于同一模块)
''   macMap: {btnName → btnName_click}  (回调函数名 = 按钮名 + "_click")
''────────────────────────────────────────────────────────────────
Private Sub BuildDefaultMaps(ByVal modName As String, _
                             ByRef outModMap As Object, _
                             ByRef outFuncMap As Object)
    Set outModMap = P_InitDic()
    Set outFuncMap = P_InitDic()
    Dim def As Object, btnName As String
    Dim i As Long
    For i = 0 To mCtrlDefs.count - 1
        Set def = mCtrlDefs(i)
        If def("Type") = "Forms.CommandButton.1" Then
            btnName = def("Name")
            outModMap(btnName) = modName
            outFuncMap(btnName) = btnName & "_click"
        End If
    Next
End Sub

' ═══════════════════════════════════════════════════════════════
'  Internal: 自包含工具函数 (替代 KCL 依赖)
' ═══════════════════════════════════════════════════════════════

Private Function P_Initlst() As Object
    Set P_Initlst = CreateObject("System.Collections.ArrayList")
End Function

Private Function P_InitDic(Optional compareMode As Long = vbBinaryCompare) As Object
    Dim dic As Object
    Set dic = CreateObject("Scripting.Dictionary")
    dic.compareMode = compareMode
    Set P_InitDic = dic
End Function

Private Function P_getRegexp() As Object
    Set P_getRegexp = CreateObject("VBScript.RegExp")
End Function

''获取 MSAPC.Apc COM 对象
Private Function P_GetApc() As Object
    Set P_GetApc = Nothing
    Dim COMObjectName$
    #If VBA7 Then
        COMObjectName = "MSAPC.Apc.7.1"
    #ElseIf VBA6 Then
        COMObjectName = "MSAPC.Apc.6.2"
    #Else
        Exit Function
    #End If
    On Error Resume Next
        Set P_GetApc = CreateObject(COMObjectName)
    On Error GoTo 0
End Function

''获取当前VBA工程文件路径
Private Function P_GetVBAprj() As String
    P_GetVBAprj = ""
    Dim oApc As Object: Set oApc = P_GetApc()
    If oApc Is Nothing Then Exit Function
    On Error Resume Next
        P_GetVBAprj = oApc.ExecutingProject.VBProject.fileName
    On Error GoTo 0
End Function

''获取指定模块中第一个过程之前的所有代码行 (声明区+注释)
Private Function P_getbf1stproc(ByVal modName As String) As String
    P_getbf1stproc = ""
    Dim oApc As Object: Set oApc = P_GetApc()
    If oApc Is Nothing Then Exit Function
    On Error Resume Next
        Dim comp As Object
        Set comp = oApc.ExecutingProject.VBProject.VBComponents(modName)
    On Error GoTo 0
    If comp Is Nothing Then Exit Function
    Dim codemod As Object: Set codemod = comp.CodeModule
    If codemod Is Nothing Then Exit Function
    Dim i As Long, procName As String, startline As Long
    Dim prockind As Long
    For i = 1 To codemod.CountOfLines
        procName = codemod.ProcOfLine(i, prockind)
        If procName <> "" Then
            startline = codemod.ProcBodyLine(procName, prockind)
            Exit For
        End If
    Next
    If startline < 1 Then Exit Function
    P_getbf1stproc = codemod.lines(1, startline)
End Function

''创建 CAT_springWD 窗体实例的工厂函数
Private Function P_CreateSpringWD() As Object
    Set P_CreateSpringWD = New CAT_springWD
End Function


