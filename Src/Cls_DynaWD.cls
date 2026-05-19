VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Cls_DynaWD"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'===========================================================================
' Cls_DynaWD
' 统一的动态UI引擎 —— 按钮生成、事件绑定、结果收集
'===========================================================================
'----以下为弹窗的标题和控件格式，写入模块头部注释即可----
'标题格式为 %Title <Caption/Text>
'控件格式为 %UI <ControlType> <ControlName> <Caption/Text><16bit_Colorcode>
'===========================================================================

' 2种数据加载模式:
'   1. getUIcfg  —— 解析模块头部 %UI / %Title 注释 (二级弹窗)
'   2. AddUIElement        —— 纯代码手动追加控件定义 (扩展)
' 3种显示模式:
'   Show       —— 模态/非模态弹窗 (替代 cls_dynaFrm.Show)
'   ShowToolbar—— 非模态工具栏     (替代 cls_dynaFrm.ShowToolbar)
'   Alert      —— 信息弹窗         (替代 cls_dynaFrm.Alert)
' 结果回收:
'   Results / btnClicked / IsCancelled
'调用方式:
'1.全手动，适合少数按钮和function
'    Dim oWD: Set oWD = KCL.new_spWD(mdlname, 1): oWD.Show
'    Select Case oWD.btnClicked
'          Case "btn_copy":    Call cpChildren
'          Case "btn_delete":  Call DeleteChildren
'          Case Else:          Exit Sub
'      End Select
'2. 自动化1
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

Private Const mdlname As String = "Cls_DynaWD"
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
'' 入口1: 从模块名称加载模块的-- %UI / %Title 注释等
Public Sub getUIcfg(ByVal modName As String)
    Dim rawCode As String
    rawCode = P_getbf1stproc(modName)
    If rawCode = "" Then Exit Sub
    mTitle = ExtractTitle(rawCode)
    Set mCtrlDefs = parseCtrldef(rawCode)
End Sub
''────────────────────────────────────────────────────────────────
'' 入口3: 纯代码手动追加一个控件定义
''        适合在运行时动态扩展UI，无需任何注释标签''
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
'' @param isModal  vbModal 或 vbModeless
'' @param ModuleMap  非模态模式: Dict{btnName -> moduleName}
'' @param MacroMap   非模态模式: Dict{btnName -> entryPoint}
''────────────────────────────────────────────────────────────────
Public Sub Show(Optional ByVal isModal As FormShowConstants = vbModal, _
                Optional ByVal ModuleMap As Object = Nothing, _
                Optional ByVal MacroMap As Object = Nothing)
    If mExecuted And isModal = vbModal Then Exit Sub
    If isModal = vbModal Then mExecuted = True
    If mCtrlDefs.count < 1 Then Exit Sub
    If Not mWD Is Nothing Then Unload mWD
    Set mWD = New CAT_springWD: mWD.Tag = ""
    Dim isbtnVert As Boolean
    If IsEmpty(mIsVertical) Then
        isbtnVert = (isModal = vbModeless)
    Else
        isbtnVert = mIsVertical
    End If
    mWD.setFrm mTitle, mCtrlDefs, isbtnVert     ' 渲染控件到窗体
    If mIsAlertMode Then
        SetupAlertMode
    Else
        BindEvents mWD, ModuleMap, MacroMap
    End If
    mWD.Show isModal
    ' 模态窗关闭后收集结果
    If isModal = vbModal And mIsAlertMode = False Then
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
'' modMap/FunMap 可选:
''   不传 → 自动按约定构建: btnName→modName, btnName→"btnName_click"
''   传入 → 使用自定义映射 (如全部指向 Action_ClickHandler)
Public Sub ShowToolbar(ByVal modName As String, _
                       Optional ByVal modMap As Object = Nothing, _
                       Optional ByVal FunMap As Object = Nothing)
    getUIcfg modName
    Me.isVertical = True
    If modMap Is Nothing Or FunMap Is Nothing Then  ' 未传映射时自动构建
        Dim autoMdl As Object, autoFunc As Object
        BuildDefaultMaps modName, autoMdl, autoFunc
        If modMap Is Nothing Then Set modMap = autoMdl
        If FunMap Is Nothing Then Set FunMap = autoFunc
    End If
    Show vbModeless, modMap, FunMap
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
Public Property Get btnClicked() As Variant
    EnsureExecuted
    btnClicked = IIf(mResultDict.Exists("btn_clicked"), mResultDict("btn_clicked"), "")
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
    If Not regEx.test(rawCode) Then Exit Function
    
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
Private Function parseCtrldef(ByVal rawCode As String) As Object
    Set parseCtrldef = P_Initlst()
    Dim lines As Variant
    lines = Split(rawCode, vbLf)
    Dim i As Long
    Dim def As Object
    For i = LBound(lines) To UBound(lines)
        Set def = ParseSingleUILine(CStr(lines(i)))
        If Not def Is Nothing Then
            parseCtrldef.Add def
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
    If Not regEx.test(sLine) Then Exit Function
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
'
'' ═══════════════════════════════════════════════════════════════
''  Internal: 菜单扫描辅助函数
'''尝试解析一个标准模块为菜单项, 失败返回 Nothing
'Private Function TryParseMenuModule(ByVal comp As Object, _
'                                    ByVal pjtPath As String, _
'                                    ByVal GroupMap As Object) As cls_MnUI
'    Set TryParseMenuModule = Nothing
'    Dim mdl As Object: Set mdl = comp.CodeModule
'    If mdl.CountOfDeclarationLines < 1 Then Exit Function
'    Dim decCode As String
'    decCode = mdl.lines(1, mdl.CountOfDeclarationLines)
'    ' 解析标签
'    Dim menuItem As New cls_MnUI
'    If Not menuItem.InitFromCode(decCode, mdl.Name, pjtPath) Then Exit Function
'    ' 验证 Group 在 GroupMap 中存在
'    Dim grpKey As Variant
'    If IsNumeric(menuItem.GroupName) Then grpKey = CLng(menuItem.GroupName) Else grpKey = menuItem.GroupName
'    If Not GroupMap.Exists(grpKey) Then Exit Function
'   ' 验证入口点存在
'    If Not ModuleMethodExists(mdl, menuItem.EntryPoint) Then
'        If ModuleMethodExists(mdl, "CATMain") Then
'            menuItem.EntryPoint = "CATMain"
'        Else
'            Exit Function
'        End If
'    End If
'    Set TryParseMenuModule = menuItem
'End Function
''检查模块中是否存在指定方法
Private Function ModuleMethodExists(ByVal mdl As Object, ByVal procName As String) As Boolean
    On Error Resume Next
    Dim line As Long
    line = mdl.ProcBodyLine(procName, 0)
    ModuleMethodExists = (line > 0)
    On Error GoTo 0
End Function
''对 ArrayList<Dict> 按模块名排序 (冒泡排序)
Private Function SortDictListByModule(ByVal lst As Object) As Object
    Dim i As Long, j As Long, temp As Object
    For i = 0 To lst.count - 1
        For j = i + 1 To lst.count - 1
            If lst(i)(TAG_MDLNAME) > lst(j)(TAG_MDLNAME) Then
                Set temp = lst(i)
                Set lst(i) = lst(j)
                Set lst(j) = temp
            End If
        Next j
    Next i
    Set SortDictListByModule = lst
End Function
' ═══════════════════════════════════════════════════════════════
'  Internal: 事件绑定
' ═══════════════════════════════════════════════════════════════
''为窗体中的所有按钮绑定事件处理
Private Sub BindEvents(frm As Object, modMap As Object, FunMap As Object)
    Dim ctl As Control, btnName As String
    Dim handler As Cls_allBTNEVT, mBTNinfo As Object
    Dim isScript As Boolean
    isScript = (Not modMap Is Nothing) And (Not FunMap Is Nothing)
    Set mEventCol = New Collection
    For Each ctl In frm.Controls
        If TypeName(ctl) = "CommandButton" Then
            btnName = ctl.Name
            Set handler = New Cls_allBTNEVT
            If isScript And Not modMap Is Nothing Then
                If modMap.Exists(btnName) And FunMap.Exists(btnName) Then
                    Set mBTNinfo = P_InitDic()
                    mBTNinfo.Add TAG_PJTPATH, mPjtPath
                    mBTNinfo.Add TAG_MDLNAME, modMap(btnName)
                    mBTNinfo.Add TAG_ENTRYPNT, FunMap(btnName)
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
' ══════════════════════════════════════════════════════════════
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
'' 等效于原 KCL.setBTNmdl + KCL.setBTNFunc,'' 默认约定:
''   modMap: {btnName → modName}       (所有按钮属于同一模块)
''   FunMap: {btnName → btnName_click}  (回调函数名 = 按钮名 + "_click")
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


''====================================================
'
''===========================================================================
'' Example_DynaWD_Usage.bas
'' Cls_DynaWD 弹窗调用示例集
'' 涵盖：所有控件类型 / 所有模态类型 / AddUIElement手动构建 / %UI注释驱动 / Alert
''===========================================================================
'
'
'
'' ' 创建方式
'' Dim oWD As New Cls_DynaWD          ' 直接 New
'' Set oWD = KCL.new_spWD("模块名")   ' 工厂方法（推荐）
'' ' 控件类型别名（AddUIElement 第一参数）
'' "Button" / "btn" / "cmd"    → CommandButton
'' "Label"  / "lbl"            → Label
'' "Text"   / "txt" / "TextBox"→ TextBox
'' "CheckBox"/"chk"            → CheckBox
'' ' 三种显示方式
'' oWD.Show vbModal              ' 阻塞，等用户点击
'' oWD.Show vbModeless           ' 非阻塞，立即返回
'' oWD.ShowToolbar "模块名"      ' 非阻塞纵向工具栏，自动绑定 btnXxx_click
'' oWD.Alert "错误内容"           ' 特殊阻塞提示框，含复制按钮
'' ' 读取结果
'' oWD.ClickedButton             ' 被点击的按钮 Name
'' oWD.IsCancelled               ' True = 未点击按钮（X 关闭）
'' oWD.Results("控件Name")       ' TextBox/CheckBox 的输入值
'
'Option Explicit
'Private Const mdlname As String = "Example_DynaWD_Usage"
'
'' ============================================================
'' [示例 1] 最简单的模态对话框 —— 仅按钮，手动 AddUIElement 构建
''          模态类型: vbModal（阻塞，等待用户点击后继续执行）
'' ============================================================
'
'Sub Demo_Modal_ButtonsOnly()
'    Dim oWD As Cls_DynaWD
'    Set oWD = New Cls_DynaWD
'
'    oWD.Title = "操作确认"
'
'    ' 添加按钮控件 (CtrlType 支持: "Button" / "btn" / "cmd")
'    oWD.AddUIElement "Button", "btn_ok", "? 确认"
'    oWD.AddUIElement "Button", "btn_cancel", "? 取消"
'
'    ' vbModal：显示后阻塞，直到用户点击按钮或关闭窗口
'    oWD.Show vbModal
'
'    ' 通过 ClickedButton 判断用户操作
'    Select Case oWD.ClickedButton
'        Case "btn_ok":     MsgBox "用户点击了 [确认]"
'        Case "btn_cancel": MsgBox "用户点击了 [取消]"
'        Case Else:         MsgBox "用户直接关闭了窗口（IsCancelled = " & oWD.IsCancelled & "）"
'    End Select
'
'    Set oWD = Nothing
'End Sub
'
'
'' ============================================================
'' [示例 2] 带所有控件类型的模态对话框
''          控件类型：Button / Label / TextBox / CheckBox
''          模态类型: vbModal
'' ============================================================
'Sub Demo_Modal_AllControlTypes()
'    Dim oWD As Cls_DynaWD
'    Set oWD = New Cls_DynaWD
'    oWD.Title = "综合控件演示"
'    oWD.AddUIElement "Label", "lbl_hint", "请填写以下信息："
'    oWD.AddUIElement "TextBox", "txt_name", "默认文本内容"
'     oWD.AddUIElement "CheckBox", "chk_option", "启用高级模式"
'    oWD.AddUIElement "Button", "btn_ok", "确  定", "#1E88E5"         ' 蓝色
'    oWD.AddUIElement "Button", "btn_cancel", "取  消", "#E53935"     ' 红色
'    ' 纵向排列（可选，不设置时 Modal 默认水平排列）
'    ' oWD.isVertical = True
'
'    oWD.Show vbModal
'
'    If oWD.IsCancelled Then
'        MsgBox "已取消或关闭"
'        Set oWD = Nothing
'        Exit Sub
'    End If
'
'    ' Results 字典包含所有控件的值（TextBox/CheckBox 的当前内容）
'    Dim res As Object
'    Set res = oWD.Results
'
'    Dim info As String
'    info = "点击按钮: " & oWD.ClickedButton & vbLf
'
'    ' 读取 TextBox 的值（键 = 控件 Name）
'    If res.Exists("txt_name") Then
'        info = info & "文本框内容: " & res("txt_name") & vbLf
'    End If
'    ' 读取 CheckBox 的值
'    If res.Exists("chk_option") Then
'        info = info & "复选框勾选: " & res("chk_option") & vbLf
'    End If
'
'    MsgBox info
'    Set oWD = Nothing
'End Sub
'
'
'' ============================================================
'' [示例 3] 非模态浮动工具栏
''          模态类型: vbModeless（非阻塞，Show 后立即返回继续执行）
''          每个按钮点击后自动调用同模块内的 btnXxx_click 过程
'' ============================================================
''
'' %UI 注释头（供 LoadFromModuleName 解析，ShowToolbar 会自动读取）：
'' %Title 动态工具栏演示
'' %UI Button btn_zoom   放大视图
'' %UI Button btn_fit    适合窗口
'' %UI Button btn_hide   隐藏选中
''
'Sub Demo_Modeless_Toolbar()
'    Dim oWD As Cls_DynaWD
'    Set oWD = New Cls_DynaWD
'
'    ' ShowToolbar 内部自动：
'    '   1. 调用 LoadFromModuleName(mdlname) 解析本模块头的 %UI 注释
'    '   2. 强制 isVertical = True（纵向排列）
'    '   3. 以 vbModeless 显示（非阻塞）
'    '   4. 将每个 btnXxx 绑定到本模块内的 btnXxx_click() 过程
'    oWD.ShowToolbar mdlname
'
'    ' Show 立即返回，此处代码继续执行
'    MsgBox "工具栏已弹出（非模态），本对话框与工具栏并存"
'
'    ' 注意：oWD 需保持作用域，否则窗口被 GC 回收而关闭
'    '       实际使用时应将 oWD 声明为模块级变量
'    Set oWD = Nothing
'End Sub
'
'' 工具栏按钮对应的响应过程（名称规则: 按钮Name + "_click"）
'Sub btn_zoom_click()
'    MsgBox "【工具栏】点击了 放大视图"
'End Sub
'Sub btn_fit_click()
'    MsgBox "【工具栏】点击了 适合窗口"
'End Sub
'Sub btn_hide_click()
'    MsgBox "【工具栏】点击了 隐藏选中"
'End Sub
'
'
'' ============================================================
'' [示例 4] 非模态 + 自定义 ModuleMap / MacroMap
''          模态类型: vbModeless
''          用途：按钮点击执行其他模块的宏，而非本模块的过程
'' ============================================================
'Sub Demo_Modeless_CustomMap()
'    Dim oWD As Cls_DynaWD
'    Set oWD = New Cls_DynaWD
'
'    oWD.Title = "跨模块工具栏"
'    oWD.isVertical = True                           ' 纵向排列
'    oWD.AddUIElement "Button", "btn_prt", "新建零件"
'    oWD.AddUIElement "Button", "btn_asm", "新建装配"
'    oWD.AddUIElement "Button", "btn_drw", "新建工程图"
'
'    ' ModuleMap: 每个按钮名 → 所在模块名
'    Dim modMap As Object
'    Set modMap = CreateObject("Scripting.Dictionary")
'    modMap("btn_prt") = "MDL_Part2Product"
'    modMap("btn_asm") = "ASM_NewBH"
'    modMap("btn_drw") = "DRW_ShtPage"
'
'    ' MacroMap: 每个按钮名 → 实际调用的过程名（入口点）
'    Dim macMap As Object
'    Set macMap = CreateObject("Scripting.Dictionary")
'    macMap("btn_prt") = "CATMain"
'    macMap("btn_asm") = "CATMain"
'    macMap("btn_drw") = "CATMain"
'
'    ' vbModeless + 自定义映射
'    oWD.Show vbModeless, modMap, macMap
'
'    MsgBox "跨模块工具栏已弹出（非模态）"
'    Set oWD = Nothing
'End Sub
'
'
'' ============================================================
'' [示例 5] 从模块 %UI 注释驱动（LoadFromModuleName）+ vbModal
''          用途：按钮由模块头注释自动解析，结果由调用方处理
'' ============================================================
''
'' 假设 OTH_Minibox.bas 头部有如下注释：
''   ' %Title 最小包围盒工具
''   ' %UI Label lbl_hint  请选择零件
''   ' %UI txt  txt_size   null
''   ' %UI Button btnOK    确认并计算
''   ' %UI Button btncancel 取消
''
'Sub Demo_Modal_LoadFromModule()
'    Dim oWD As Cls_DynaWD
'    Set oWD = New Cls_DynaWD
'
'    ' 从指定模块名解析 %UI 和 %Title 注释，自动填充控件定义
'    oWD.LoadFromModuleName "OTH_Minibox"
'
'    ' 模态显示，等待用户操作
'    oWD.Show vbModal
'
'    If oWD.IsCancelled Then
'        Set oWD = Nothing
'        Exit Sub
'    End If
'
'    ' 根据点击的按钮分支处理
'    Select Case oWD.ClickedButton
'        Case "btnOK"
'            Dim inputVal As String
'            If oWD.Results.Exists("txt_size") Then
'                inputVal = oWD.Results("txt_size")
'            End If
'            MsgBox "用户输入尺寸: " & inputVal
'        Case "btncancel"
'            MsgBox "操作已取消"
'    End Select
'
'    Set oWD = Nothing
'End Sub
'
'
'' ============================================================
'' [示例 6] Alert 模式 —— 错误/信息提示框（含"复制内容"按钮）
''          这是特殊的模态对话框，专门用于展示长文本并允许复制
''          模态类型: vbModal（内部固定，Alert 方法自动处理）
'' ============================================================
'Sub Demo_Alert()
'    Dim oWD As Cls_DynaWD
'    Set oWD = New Cls_DynaWD
'
'    ' 必须先至少添加一个控件（供窗口构建 TextBox 用）
'    oWD.AddUIElement "TextBox", "txt_msg", ""
'
'    ' Alert 会将 Content 填入 TextBox，并添加"复制内容"按钮
'    ' 用户点击"复制内容"后，文本自动进入剪贴板，窗口关闭
'    Dim errInfo As String
'    errInfo = "错误详情：" & vbLf & _
'              "  行号: 42" & vbLf & _
'              "  模块: ASM_ChildMng" & vbLf & _
'              "  描述: 找不到根产品，请确认已打开装配文档"
'
'    oWD.Alert errInfo
'
'    ' Alert 关闭后继续执行
'    MsgBox "Alert 已关闭，用户已看到错误信息"
'    Set oWD = Nothing
'End Sub
'
'
'' ============================================================
'' [示例 7] 综合场景：多步骤向导（两个连续的模态对话框）
''          演示如何将 Cls_DynaWD 串联使用
'' ============================================================
'Sub Demo_MultiStep_Wizard()
'
'    ' --- Step 1: 选择操作类型 ---
'    Dim step1 As Cls_DynaWD
'    Set step1 = New Cls_DynaWD
'    step1.Title = "Step 1/2 — 选择操作"
'    step1.AddUIElement "Label", "lbl_s1", "请选择要执行的操作类型："
'    step1.AddUIElement "Button", "btn_copy", "复制子零件", "#43A047"
'    step1.AddUIElement "Button", "btn_delete", "删除子零件", "#E53935"
'    step1.AddUIElement "Button", "btn_exit", "退出", "#9E9E9E"
'    step1.Show vbModal
'
'    Dim action As String
'    action = step1.ClickedButton
'    Set step1 = Nothing
'
'    If action = "btn_exit" Or action = "" Then Exit Sub
'
'    ' --- Step 2: 输入参数 ---
'    Dim step2 As Cls_DynaWD
'    Set step2 = New Cls_DynaWD
'    step2.Title = "Step 2/2 — 配置参数"
'    step2.AddUIElement "Label", "lbl_s2", "操作: " & IIf(action = "btn_copy", "复制", "删除")
'    step2.AddUIElement "TextBox", "txt_count", "1"
'    step2.AddUIElement "CheckBox", "chk_log", "记录操作日志"
'    step2.AddUIElement "Button", "btn_run", "执行", "#1E88E5"
'    step2.AddUIElement "Button", "btn_back", "返回"
'    step2.Show vbModal
'
'    If step2.IsCancelled Or step2.ClickedButton = "btn_back" Then
'        MsgBox "已返回或取消"
'    Else
'        Dim cnt As String: cnt = ""
'        If step2.Results.Exists("txt_count") Then cnt = step2.Results("txt_count")
'        Dim doLog As Boolean: doLog = False
'        If step2.Results.Exists("chk_log") Then doLog = CBool(step2.Results("chk_log"))
'        MsgBox "执行操作: " & action & vbLf & "数量: " & cnt & vbLf & "记录日志: " & doLog
'    End If
'
'    Set step2 = Nothing
'End Sub
'
'
'' ============================================================
'' [示例 8] 使用 KCL.new_spWD 工厂方法（推荐写法）
''          等效于 New Cls_DynaWD + LoadFromModuleName
'' ============================================================
'Sub Demo_KCL_Factory()
'    ' 写法A：直接从模块名创建（最简）
'    Dim oWD As Object
'    Set oWD = KCL.new_spWD("OTH_Minibox")
'    oWD.Show vbModal
'
'    If Not oWD.IsCancelled Then
'        MsgBox "点击了: " & oWD.ClickedButton
'    End If
'    Set oWD = Nothing
'
'    ' 写法B：空白创建后手动添加控件（链式风格）
'    Dim oWD2 As Object
'    Set oWD2 = KCL.new_spWD()      ' 不传模块名 = 空白画布
'    oWD2.Title = "工厂方法示例"
'    oWD2.AddUIElement "Button", "btn_yes", "是"
'    oWD2.AddUIElement "Button", "btn_no", "否"
'    oWD2.Show vbModal
'    MsgBox "结果: " & oWD2.ClickedButton
'    Set oWD2 = Nothing
'End Sub
'
'
