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
Private mExecuted
Private mdict
Private pobj As Object
Private WD

Private Const mdlname As String = "cls_dynaFrm"
Private Sub Class_Initialize()
        mExecuted = False
        Set mdict = CreateObject("Scripting.Dictionary")
        Set Me.mdl = Nothing
         mExecuted = False
End Sub
Private Sub Class_Terminate()
On Error Resume Next
 Set mdict = Nothing
 Set Me.mdl = Nothing
 Set WD = Nothing
 Set Me.Res = Nothing
  mExecuted = False
        Error.Clear
        On Error GoTo 0
End Sub
Public Property Get mdl()
   Set mdl = pobj
End Property

Public Property Set mdl(value)
   Set pobj = value
End Property

Private Function Execute()
  If mExecuted = True Then Exit Function
    mExecuted = True
    Dim DecCnt, DecCode, i, procname, startline
    Dim prockind As Long
    If Me.mdl Is Nothing Then Exit Function
   Dim codeMod: Set codeMod = mdl.CodeModule
    For i = 1 To codeMod.CountOfLines
        procname = codeMod.ProcOfLine(i, prockind)
        If procname <> "" Then startline = codeMod.ProcBodyLine(procname, prockind)  ' procKind是枚举:sub/function/property
            If startline > 1 Then Exit For
    Next
    If startline < 1 Then Exit Function
    
    DecCode = codeMod.Lines(1, startline)
    Dim clscfg: Set clscfg = ParseCts(DecCode)
    Dim ttl: ttl = ParseTitle(DecCode)
    
   Set WD = New springWD: WD.Tag = ""
    Call WD.setFrm(ttl, clscfg)
    ' --- 3. bind event handler ---
    Dim eventCol As New collection
    Dim ctl As Control
'Dim evtHandler As Cls_DynaFrm_EventHandler
  Dim evtHandler As Cls_btEVT
    For Each ctl In WD.Controls
        If TypeName(ctl) = "CommandButton" Then
           ' Set evtHandler = New Cls_DynaFrm_EventHandler
            
          Set evtHandler = New Cls_btEVT
            Set evtHandler.ControlBtn = ctl
            evtHandler.ControlID = ctl.Name
            Set evtHandler.ParentFrm = WD
            eventCol.Add evtHandler
        End If
    Next
    ' --- 4. show form (modal) ---
    WD.Show vbModal  'vbModal vbModeless
    ' --- 5. collect result ---
    If WD.Tag = "" Then
        WD.Tag = "UserClosed"
        Set WD = Nothing
        Exit Function
    End If
    
    mdict.Add "btn_clicked", WD.Tag
    
    For Each ctl In WD.Controls
        On Error Resume Next
        If TypeName(ctl) = "CheckBox" Then
             mdict.Add ctl.Name, ctl.value
        ElseIf TypeName(ctl) = "TextBox" Then
             mdict.Add ctl.Name, ctl.Text
        End If
        On Error GoTo 0
    Next
    Unload WD
        On Error Resume Next
        Set WD = Nothing
        Error.Clear
        On Error GoTo 0
End Function

Public Property Get Res() As Object
   Call Execute:   Set Res = mdict
End Property
Public Property Get BtnClicked(Optional key = Null) As Variant
   Call Execute
        On Error Resume Next
            Dim btnDic: Set btnDic = Nothing
            Set btnDic = mdict
        On Error GoTo 0
        If btnDic Is Nothing Then Exit Property ' Return Null if not executed/failed
       
        If btnDic.Exists("btn_clicked") = True Then
                    If key <> Null Then
                                    If btnDic("btn_clicked") = LCase(key) Then
                                         BtnClicked = True
                                    Else
                                         BtnClicked = False
                                    End If
                          Exit Property
                    Else
                        BtnClicked = btnDic("btn_clicked")
                    End If
        Else
                BtnClicked = ""
        End If
End Property
Public Property Get IsCancelled() As Boolean
   Call Execute
    If mdict Is Nothing Then
        IsCancelled = True
        Exit Property
    End If
    If mdict.Exists("btn_clicked") Then
        If mdict("btn_clicked") = "UserClosed" Then
            IsCancelled = True
        Else
            IsCancelled = False
        End If
    Else
        IsCancelled = True
    End If
End Property

' getiType Helper
Function getiType(ctrltypename As String) As String
    ' Static because it's constant map
    Static mapping As Object
    If mapping Is Nothing Then
        Set mapping = CreateObject("Scripting.Dictionary")
        ' Key = alias
        ' Value = Control ID
        mapping.Add "commandbutton,button,cmd,cbt,btn", "Forms.CommandButton.1"
        mapping.Add "textbox,text,txt", "Forms.TextBox.1"
        mapping.Add "label,lbl", "Forms.Label.1"
        mapping.Add "checkbox,check,chk", "Forms.CheckBox.1"
        mapping.Add "optionbutton,option,opt", "Forms.OptionButton.1"
        mapping.Add "listbox,list,lst", "Forms.ListBox.1"
        mapping.Add "combobox,combo,cmb", "Forms.ComboBox.1"
        mapping.Add "multipage,multipages,mpg", "Forms.MultiPage.1"
    End If
    Dim key As Variant
    For Each key In mapping.keys
        If ContainsAny(ctrltypename, CStr(key)) Then
            getiType = mapping(key)
            Exit Function
        End If
    Next key
    getiType = "Forms.Label.1"
End Function

' Helper
Private Function ContainsAny(ByVal mainString As String, ByVal substringsList As String) As Boolean
    Dim substrings() As String
    Dim substring As Variant
    If substringsList = "" Then
        ContainsAny = False
        Exit Function
    End If
    substrings = Split(substringsList, ",")
    ContainsAny = False
    For Each substring In substrings
        If InStr(1, mainString, VBA.Trim(substring), vbTextCompare) > 0 Then
            ContainsAny = True
            Exit Function
        End If
    Next substring
End Function

Private Function ExistsKey(ByVal txt As String, ByVal key As String) As Boolean
    ExistsKey = IIf(VBA.InStr(VBA.LCase(txt), VBA.LCase(key)) > 0, True, False)
End Function

Private Function ParseCts(ByVal code As String) As Object
    Dim regEx As Object
    Dim matches As Object
    Dim match As Object
    Dim Cls_property As Object ' Scripting.Dictionary
    Set regEx = CreateObject("VBScript.RegExp")
    With regEx
        .Global = True
        .MultiLine = True
        ' Pattern: ' %UI <ControlType> <ControlName> <Caption/Text>
       .Pattern = "^\s*'\s*%UI\s+(\w+)\s+(\w+)\s+(.*)$"
    End With
    Dim lst: Set lst = KCL.InitLst
    If regEx.TEST(code) Then
        Set matches = regEx.Execute(code)
        For Each match In matches
            Set Cls_property = KCL.InitDic
            Cls_property.Add "Type", getiType(match.SubMatches(0))
            Cls_property.Add "Name", match.SubMatches(1)
            Cls_property.Add "Caption", VBA.Trim(match.SubMatches(2))
'            If Not ctrlcoll.Contains(Cls_property("Name")) Then
               lst.Add Cls_property
'            End If
        Next
    End If
    Set ParseCts = lst
End Function
Private Function ParseTitle(ByVal code As String) As String
        Dim regEx As Object
        Dim matches As Object
        Dim match As Object
        Dim itl
        Set regEx = KCL.getRegexp
    With regEx
        .Global = True
        .MultiLine = True
        ' 格式: ' %UI <Title> <Caption/Text>
           .Pattern = "^\s*'\s*%Title\s+(.*)$"
    End With
    itl = "请配置你要如何执行"
              If regEx.TEST(code) Then
                Set matches = regEx.Execute(code)
                For Each match In matches
                 itl = match.SubMatches(0)
                Next
              End If
    ParseTitle = itl
End Function


