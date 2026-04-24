VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cls_MnUI"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Option Explicit
' --- Private Variables ---
Private mTitle As String
Private mGroupName As Variant ' Changed to Variant to hold Long or String
Private mModulePath As String
Private mModuleName As String
Private mEntryPoint As String
Private mParams As Variant
Private mExtras As Object
' --- Constants ---
Private Const TAG_S = "{"
Private Const TAG_D = ":"
Private Const TAG_E = "}"
Private Const TAG_GROUP = "gp"
Private Const TAG_ENTRYPNT = "ep"
Private Const TAG_ENTRY_DEF = "CATMain"
Private Const TAG_PJTPATH = "pjt_path"
Private Const TAG_MDLNAME = "mdl_name"
' --- Properties ---
Private Const mdlname As String = "cls_MnUI"
Public Property Get Title() As String
    Title = mTitle
End Property
Public Property Let Title(ByVal Value As String)
    mTitle = Value
End Property
Public Property Get GroupName() As Variant
    GroupName = mGroupName
End Property
Public Property Let GroupName(ByVal Value As Variant)
    mGroupName = Value
End Property
Public Property Get ModulePath() As String
    ModulePath = mModulePath
End Property
Public Property Let ModulePath(ByVal Value As String)
    mModulePath = Value
End Property
Public Property Get ModuleName() As String
    ModuleName = mModuleName
End Property
Public Property Let ModuleName(ByVal Value As String)
    mModuleName = Value
End Property
Public Property Get EntryPoint() As String
    If mEntryPoint = "" Then EntryPoint = TAG_ENTRY_DEF Else EntryPoint = mEntryPoint
End Property
Public Property Let EntryPoint(ByVal Value As String)
    mEntryPoint = Value
End Property
Public Property Get params() As Variant
    params = mParams
End Property
Public Property Let params(ByVal Value As Variant)
    mParams = Value
End Property
' --- Initialization ---
Private Sub Class_Initialize()
    mParams = Array()
    Set mExtras = CreateObject("Scripting.Dictionary")
    mExtras.compareMode = vbTextCompare
End Sub
' --- Methods ---
Public Sub Run()
    Dim ss As Object
    Set ss = CATIA.SystemService
    Dim pjtPath As String
    If mModulePath = "" Then
         On Error Resume Next
         pjtPath = GetApc().ExecutingProject.DisplayName
         On Error GoTo 0
    Else
        pjtPath = mModulePath
    End If
    On Error Resume Next
    Call ss.ExecuteScript(pjtPath, _
                          catScriptLibraryTypeVBAProject, _
                          mModuleName, _
                          Me.EntryPoint, _
                          mParams)
    
    If Err.Number <> 0 Then
        MsgBox "Error executing macro: " & mModuleName & "." & Me.EntryPoint & vbCrLf & _
               "Error: " & Err.Description, vbCritical
        Err.Clear
    End If
    On Error GoTo 0
End Sub

Private Function GetApc() As Object
    Set GetApc = Nothing
    Dim COMObjectName$
    #If VBA7 Then
        COMObjectName = "MSAPC.Apc.7.1"
    #ElseIf VBA6 Then
        COMObjectName = "MSAPC.Apc.6.2"
    #Else
        Exit Function
    #End If
    On Error Resume Next
        Set GetApc = CreateObject(COMObjectName)
    On Error GoTo 0
End Function

Public Function InitFromCode(ByVal decCode As String, ByVal mdlname As String, ByVal currentPjtPath As String) As Boolean
    InitFromCode = False
    mModuleName = mdlname
    mModulePath = currentPjtPath
    mExtras.RemoveAll
    ' Quick check for Group tag to avoid heavy regex
    If InStr(1, decCode, TAG_GROUP, vbTextCompare) = 0 Then Exit Function
    Dim Reg As Object: Set Reg = CreateObject("VBScript.RegExp")
    With Reg
        .Pattern = TAG_S & "(.*?)" & TAG_D & "(.*?)" & TAG_E
        .Global = True
    End With
    Dim matches As Object: Set matches = Reg.Execute(decCode)
    Set Reg = Nothing
    If matches.count < 1 Then Exit Function
    Dim match As Object, SubMatchs As Object
    Dim KEY As String, val As Variant ' Variant to allow Long conversion
    Dim hasGroup As Boolean
    For Each match In matches
        Set SubMatchs = match.SubMatches
        If SubMatchs.count >= 2 Then
            KEY = Trim(Replace(SubMatchs(0), """", ""))
            val = Trim(Replace(SubMatchs(1), """", ""))
            ' Try Convert to Long if it looks numeric (Strictly mimic m0_dataMenu line 87)
            If IsNumeric(val) Then val = CLng(val)
            ' 1. Store in Extras (using updated VAL type)
            If mExtras.Exists(KEY) Then mExtras(KEY) = val Else mExtras.Add KEY, val
            ' 2. Map to properties
            Select Case LCase(KEY)
                Case LCase(TAG_GROUP)
                    mGroupName = val
                    If IsNumeric(val) Then mTitle = CStr(val) ' Title is string
                    hasGroup = True
                Case LCase(TAG_ENTRYPNT)
                    mEntryPoint = val
            End Select
        End If
    Next
    
    If hasGroup Then InitFromCode = True
End Function

Public Function ToDictionary() As Object
    Dim dic As Object
    Set dic = CreateObject("Scripting.Dictionary")
    dic.compareMode = vbTextCompare
    Dim k
    For Each k In mExtras.keys
        dic.Add k, mExtras(k)
    Next
    If dic.Exists(TAG_MDLNAME) Then dic(TAG_MDLNAME) = mModuleName Else dic.Add TAG_MDLNAME, mModuleName
    If dic.Exists(TAG_PJTPATH) Then dic(TAG_PJTPATH) = mModulePath Else dic.Add TAG_PJTPATH, mModulePath
    ' EP might default to CATMain if empty, update it
    If dic.Exists(TAG_ENTRYPNT) Then dic(TAG_ENTRYPNT) = Me.EntryPoint Else dic.Add TAG_ENTRYPNT, Me.EntryPoint
    Set ToDictionary = dic
End Function






