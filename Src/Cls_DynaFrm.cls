VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Cls_DynaFrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Public WithEvents ControlBtn As MSForms.CommandButton
Attribute ControlBtn.VB_VarHelpID = -1
Public ParentFrm As Object ' 指向当前的 UserForm
Public ControlID As String ' 用于识别是哪个按钮（如 "btnOK" 或 "btnCancel"）
Private mdict
Private mExecuted As Boolean
Private Sub Class_Initialize()
    ' Initialization is now lazy to prevent recursion
End Sub
Private Sub Execute()
    If mExecuted Then Exit Sub
    mExecuted = True
    Dim DecCnt, DecCode
    Dim Apc As Object: Set Apc = KCL.GetApc()
    Dim ExecPjt As Object: Set ExecPjt = Apc.ExecutingProject
    Dim mdl As Object
    On Error Resume Next
    Set mdl = ExecPjt.VBProject.VBE.Activecodepane.codemodule
    On Error GoTo 0
    
    If mdl Is Nothing Then Exit Sub
    DecCnt = mdl.CountOfDeclarationLines
    If DecCnt < 1 Then Exit Sub
    DecCode = mdl.Lines(1, DecCnt)
    
    Dim clscfg: Set clscfg = ParseCts(DecCode)
    Dim ttl: ttl = ParseTitle(DecCode)
    Dim WD As springWD
    Set WD = New springWD
    WD.Tag = ""
    Call WD.setFrm(ttl, clscfg)
    
    ' --- 3. 绑定按钮事件 ---
    Dim eventCol As New collection
    Dim ctl As Control
    Dim evtHandler As Cls_DynaFrm
    For Each ctl In WD.Controls
        If TypeName(ctl) = "CommandButton" Then
            Set evtHandler = New Cls_DynaFrm
            Set evtHandler.ControlBtn = ctl
            Set evtHandler.ParentFrm = WD
            evtHandler.ControlID = ctl.Name
            eventCol.Add evtHandler
        End If
    Next
    
    ' --- 4. 显示 (模态) ---
    WD.Show vbModal
    
    ' --- 5. 处理结果 ---
    Set mdict = CreateObject("Scripting.Dictionary")
    
    ' 如果 WD.Tag 为空（即不是通过点击按钮关闭，也不是由 QueryClose 显式设置），
    ' 则可能是用户关闭了窗体。虽然 springWD 已经处理了 QueryClose，但这里做个保底。
    Dim clicked As String
    clicked = WD.Tag
    If clicked = "" Then clicked = "UserClosed"
    mdict.Add "btn_clicked", clicked
    
    For Each ctl In WD.Controls
        On Error Resume Next
        If TypeName(ctl) = "CheckBox" Then
             mdict.Add ctl.Name, ctl.value
        ElseIf TypeName(ctl) = "TextBox" Then
             mdict.Add ctl.Name, ctl.Text
        End If
        On Error GoTo 0
    Next
    
    ' --- 6. 卸载 ---
    Unload WD
End Sub

Public Property Get Res() As Object
   Execute
   Set Res = mdict
End Property
Public Property Get BtnClicked(Optional key = Null) As Variant
        Execute
        On Error Resume Next
            Dim btnDic: Set btnDic = Nothing
            Set btnDic = mdict
        On Error GoTo 0
        If btnDic Is Nothing Then Exit Property ' Return Null if not executed/failed
       
        If btnDic.Exists("btn_clicked") = True Then  '若有按钮按下
              If key <> Null Then    '若有key先比较key
                    If btnDic("btn_clicked") = LCase(key) Then
                         BtnClicked = True
                    Else
                         BtnClicked = False
                    End If
                 Exit Property
                Else    '若无key则输出对应按钮名字
                  BtnClicked = btnDic("btn_clicked")
                End If
        Else  '若无按钮按下
                BtnClicked = ""
        End If
        
        
End Property

Public Property Get IsCancelled() As Boolean
    Execute
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

Private Sub ControlBtn_Click()
    ParentFrm.Tag = ControlID  ' 1. 将点击的按钮 ID 存入窗体的 Tag 属性，作为返回值标识
    ParentFrm.Hide   ' 2. 隐藏窗体（注意：不是 Unload），
End Sub


' getiType 函数（字典驱动版）
Function getiType(ctrltypename As String) As String
    ' 使用 Static 关键字，字典在多次函数调用之间只会创建一次，提高性能
    Static mapping As Object
    ' 如果是第一次调用，则创建并填充字典
    If mapping Is Nothing Then
        Set mapping = CreateObject("Scripting.Dictionary")
        ' Key = 逗号分隔的别名
        ' Value = 控件类型ID
        mapping.Add "commandbutton,button,cmd,cbt,btn", "Forms.CommandButton.1"
        mapping.Add "textbox,text,txt", "Forms.TextBox.1"
        mapping.Add "label,lbl", "Forms.Label.1"
        mapping.Add "checkbox,check,chk", "Forms.CheckBox.1"
        mapping.Add "optionbutton,option,opt", "Forms.OptionButton.1"
        mapping.Add "listbox,list,lst", "Forms.ListBox.1"
        mapping.Add "combobox,combo,cmb", "Forms.ComboBox.1"
        mapping.Add "multipage,multipages,mpg", "Forms.MultiPage.1"
    End If
    ' 遍历字典中的所有别名列表（Key）
    Dim key As Variant
    For Each key In mapping.keys
        ' 如果输入名称包含了任何一个别名
        If ContainsAny(ctrltypename, CStr(key)) Then
            ' 则返回对应的控件类型ID并退出函数
            getiType = mapping(key)
            Exit Function
        End If
    Next key
    ' 如果遍历完字典都没有找到匹配项，则返回默认类型
    getiType = "Forms.Label.1"
End Function

' 这个辅助函数保持不变
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
        ' 格式: ' %UI <ControlType> <ControlName> <Caption/Text>
        .Pattern = "^\s*'\s*%UI\s+(\w+)\s+(\w+)\s+(.*)$"
    End With
    Dim lst: Set lst = KCL.InitLst
    If regEx.test(code) Then
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
              If regEx.test(code) Then
                Set matches = regEx.Execute(code)
                For Each match In matches
                 itl = match.SubMatches(0)
                Next
              End If
    ParseTitle = itl
End Function





