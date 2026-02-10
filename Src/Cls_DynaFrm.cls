VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cls_dynaFrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' --- State ---
Private mForm, mdict, WD As Object
Private mExecuted As Boolean
Private mCfgStr As String
Private mEventCol As Collection
' Clipboard API
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

Private mIsVertical As Variant ' Empty=Auto, True=Vert, False=Horiz
Private mPassBtn As Boolean
Private mIsAlertMode As Boolean
Private mAlertMsg As String
Private WithEvents mCopyBtn As MSForms.CommandButton
Attribute mCopyBtn.VB_VarHelpID = -1

' --- Constants ---
Private Const TAG_ENTRYPNT = "ep"           ' 入口点标签
Private Const TAG_PJTPATH = "pjt_path"      ' 项目路径标签
Private Const TAG_MDLNAME = "mdl_name"      ' 模块名称标签
Private Const TAG_PASS_BTN = "springwdBTN"
Private pjt As String
' --- Lifecycle ---
Private Const mdlname As String = "cls_dynaFrm"
Private Sub Class_Initialize()
    mExecuted = False
    Set mdict = CreateObject("Scripting.Dictionary")
    Set mEventCol = New Collection
    pjt = KCL.GetVBAprj
    mIsVertical = Empty
End Sub
Private Sub Class_Terminate()
    On Error Resume Next
    If Not WD Is Nothing Then Unload WD
    Set mdict = Nothing
    Set WD = Nothing
    Set mForm = Nothing
    Set mEventCol = Nothing
    mExecuted = False
    Set g_frm = Nothing
    pjt = ""
    mCfgStr = ""
    mIsAlertMode = False
    mAlertMsg = ""
    Error.Clear
End Sub

' --- Public API ---
Public Sub Init(ByVal mdlname As String)
    mCfgStr = getbf1stproc(mdlname)
End Sub
Public Property Get res() As Object
EnsureExecuted: Set res = mdict
End Property
Public Property Get BtnClicked() As Variant
EnsureExecuted: BtnClicked = IIf(mdict.Exists("btn_clicked"), mdict("btn_clicked"), "")
End Property
Public Property Get IsCancelled() As Boolean
    EnsureExecuted
    IsCancelled = IIf(mdict.Exists("btn_clicked") And mdict("btn_clicked") <> "UserClosed", False, True)
End Property

Public Property Let isVertical(ByVal val As Boolean)
    mIsVertical = val
End Property

Public Property Let PassButtonName(ByVal val As Boolean)
    mPassBtn = val
End Property

Public Sub Alert(ByVal Content As String)
    mIsAlertMode = True
    mAlertMsg = Content
    mExecuted = False ' Auto-Reset for Alert
    Show vbModal
End Sub

Public Sub Show(Optional ByVal iShowMode As FormShowConstants = vbModal, _
                Optional ByVal ModuleMap As Object = Nothing, _
                Optional ByVal MacroMap As Object = Nothing)
    If mExecuted And iShowMode = vbModal Then Exit Sub
    If iShowMode = vbModal Then mExecuted = True
    If mCfgStr = "" Then Exit Sub
    Dim ctrlCfg As Object: Set ctrlCfg = KCL.ParseUIConfig(mCfgStr)
    Dim ttl As String: ttl = KCL.ParseUITitle(mCfgStr)
    If Not WD Is Nothing Then Unload WD
    Set WD = New CAT_springWD: WD.Tag = ""
    Dim btnvert As Boolean
    If IsEmpty(mIsVertical) Then
        btnvert = (iShowMode = vbModeless) ' 自动挡：非模态则竖排
    Else
        btnvert = mIsVertical         ' 手动挡
    End If
    WD.setFrm ttl, ctrlCfg, btnvert
    
    If mIsAlertMode Then
        Dim ctl As Control
        WD.Tag = ""
        For Each ctl In WD.Controls
            If TypeName(ctl) = "CommandButton" Then
                ctl.Visible = False
            ElseIf TypeName(ctl) = "TextBox" Then
                ctl.Value = mAlertMsg
            End If
        Next
        
        ' Dynamic Copy Button
        Dim btnCopy As MSForms.CommandButton
        Set btnCopy = WD.Controls.Add("Forms.CommandButton.1", "btn_Copy", True)
        With btnCopy
            .Caption = "复制并关闭"
            .Width = 80
            .Height = 24
            .Left = (WD.InsideWidth - .Width) / 2
            .Top = WD.InsideHeight - .Height - 6
        End With
        Set mCopyBtn = btnCopy
        
    Else
        BindEvents WD, ModuleMap, MacroMap  '为窗口的按钮绑定事件
    End If
    
    WD.Show iShowMode
If iShowMode = vbModal And mIsAlertMode = False Then
       If WD.Tag = "" Then
            WD.Tag = "UserClosed": Unload WD: Set WD = Nothing
            Exit Sub
        End If
        If mdict Is Nothing Then Set mdict = KCL.InitDic
            mdict.Add "btn_clicked", WD.Tag
           WD.Hide: Unload WD: Set WD = Nothing
End If
End Sub
Public Sub ShowToolbar(ByVal modName As String, _
                       ByVal modMap As Object, _
                       ByVal macMap As Object)
    mCfgStr = KCL.getbf1stproc(modName)
    Me.isVertical = True
    Show vbModeless, modMap, macMap
End Sub
' --- Internal Logic ---
Private Sub BindEvents(frm As Object, modMap As Object, macMap As Object)
    Dim ctl As Control, btnName As String
    Dim handler As Cls_btEVT, mBTNinfo As Object
    Dim isScript As Boolean
    isScript = (Not modMap Is Nothing) And (Not macMap Is Nothing)
    
    Set mEventCol = New Collection
    For Each ctl In frm.Controls
        If TypeName(ctl) = "CommandButton" Then
            btnName = ctl.name
            Set handler = New Cls_btEVT
                If isScript And Not modMap Is Nothing Then  '非模态按钮  对应cls_btEVT set_ButtonEvent
                    If modMap.Exists(btnName) And macMap.Exists(btnName) Then
                        Set mBTNinfo = KCL.InitDic
                        mBTNinfo.Add TAG_PJTPATH, pjt
                        mBTNinfo.Add TAG_MDLNAME, modMap(btnName)
                        mBTNinfo.Add TAG_ENTRYPNT, macMap(btnName)
                        ' Add PassButtonName flag if enabled
                        If mPassBtn Then mBTNinfo.Add TAG_PASS_BTN, True
                        handler.set_ButtonEvent ctl, mBTNinfo, frm, True ' False
                    End If
                Else          '模态按钮 对应 button Clicked和Res对应结果
                    Set handler.ControlBTN = ctl
                    handler.ControlID = btnName
                    Set handler.ParentFrm = frm
                    Set handler.ResultDict = mdict
                End If
            mEventCol.Add handler
        End If
    Next
End Sub

Private Sub EnsureExecuted()
    If Not mExecuted And mCfgStr <> "" Then Show vbModal
End Sub

Private Sub mCopyBtn_Click()
    On Error Resume Next
    Dim hGlobal As LongPtr
    Dim lpGlobal As LongPtr
    Dim sText As String
    sText = mAlertMsg & vbNullChar
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
    Unload WD
'    mIsAlertMode = False

End Sub







