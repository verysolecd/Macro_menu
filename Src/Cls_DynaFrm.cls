VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cls_dynaFrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
' --- State ---
Private mForm, mdict, WD As Object
Private mExecuted As Boolean
Private mConfigStr As String
Private mEventCol As collection
' --- Constants ---
Private Const TAG_ENTRYPNT = "ep"           ' 入口点标签
Private Const TAG_PJTPATH = "pjt_path"      ' 项目路径标签
Private Const TAG_MDLNAME = "mdl_name"      ' 模块名称标签
Private pjt As String
' --- Lifecycle ---
Private Const mdlname As String = "cls_dynaFrm"
Private Sub Class_Initialize()
    mExecuted = False
    Set mdict = CreateObject("Scripting.Dictionary")
    Set mEventCol = New collection
    pjt = KCL.GetVBAprj
End Sub
Private Sub Class_Terminate()
    On Error Resume Next
    If Not WD Is Nothing Then Unload WD
    Set mdict = Nothing
    Set WD = Nothing
    Set mForm = Nothing
    Set mEventCol = Nothing
    mExecuted = False
    Set g_Frm = Nothing
    pjt = ""
    Error.Clear
End Sub
' --- Public API ---
Public Sub Init(ByVal ConfigStr As String)
    mConfigStr = ConfigStr
End Sub
Public Sub Show(Optional ByVal Mode As FormShowConstants = vbModal, _
                Optional ByVal ModuleMap As Object = Nothing, _
                Optional ByVal MacroMap As Object = Nothing)
    If mExecuted And Mode = vbModal Then Exit Sub
    If Mode = vbModal Then mExecuted = True
    If mConfigStr = "" Then Exit Sub
    Dim ctrlCfg As Object: Set ctrlCfg = KCL.ParseUIConfig(mConfigStr)
    Dim ttl As String: ttl = KCL.ParseUITitle(mConfigStr)
    If Not WD Is Nothing Then Unload WD
    Set WD = New springWD: WD.Tag = ""
    Dim btnvert:  btnvert = False
    If Mode = vbModeless Then btnvert = True
        WD.setFrm ttl, ctrlCfg, btnvert
    BindEvents WD, ModuleMap, MacroMap  '为窗口的按钮绑定事件
    WD.Show Mode
If Mode = vbModal Then
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
    mConfigStr = KCL.getbf1stproc(modName)
    Show vbModeless, modMap, macMap
End Sub
' --- Internal Logic ---
Private Sub BindEvents(frm As Object, modMap As Object, macMap As Object)
    Dim ctl As Control, btnName As String
    Dim handler As Cls_btEVT, mBTNinfo As Object
    Dim isScript As Boolean
    isScript = (Not modMap Is Nothing) And (Not macMap Is Nothing)
    Set mEventCol = New collection
    For Each ctl In frm.Controls
        If TypeName(ctl) = "CommandButton" Then
            btnName = ctl.Name
            Set handler = New Cls_btEVT
                If isScript And Not modMap Is Nothing Then  '非模态按钮  对应cls_btEVT set_ButtonEvent
                    If modMap.Exists(btnName) And macMap.Exists(btnName) Then
                        Set mBTNinfo = KCL.InitDic
                        mBTNinfo.Add TAG_PJTPATH, pjt
                        mBTNinfo.Add TAG_MDLNAME, modMap(btnName)
                        mBTNinfo.Add TAG_ENTRYPNT, macMap(btnName)
                        handler.set_ButtonEvent ctl, mBTNinfo, frm, True
                    End If
                Else                 '模态按钮
                    Set handler.ControlBTN = ctl
                    handler.ControlID = btnName
                    Set handler.ParentFrm = frm
                    Set handler.ResultDict = mdict
                End If
            mEventCol.Add handler
        End If
    Next
End Sub
' --- Properties ---
Public Property Get Values() As Object
EnsureExecuted: Set Values = mdict
End Property
Public Property Get Res() As Object
EnsureExecuted: Set Res = mdict
End Property
Public Property Get BtnClicked() As Variant
EnsureExecuted: BtnClicked = IIf(mdict.Exists("btn_clicked"), mdict("btn_clicked"), "")
End Property
Public Property Get IsCancelled() As Boolean
    EnsureExecuted
    IsCancelled = IIf(mdict.Exists("btn_clicked") And mdict("btn_clicked") <> "UserClosed", False, True)
End Property
Private Sub EnsureExecuted()
    If Not mExecuted And mConfigStr <> "" Then Show vbModal
End Sub

