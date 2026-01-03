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

' 声明带事件的按钮对象
Public WithEvents ControlBtn As MSForms.CommandButton
Attribute ControlBtn.VB_VarHelpID = -1
Public ParentFrm As Object ' 指向当前的 UserForm
Public ControlID As String ' 用于识别是哪个按钮（如 "btnOK" 或 "btnCancel"）
Private mRes As Object   '参数本身
Private Sub Class_Initialize()


End Sub
Public Property Get Res()
   Set Res = getUIres
End Property
Private Sub ControlBtn_Click()
    ParentFrm.Tag = ControlID  ' 1. 将点击的按钮 ID 存入窗体的 Tag 属性，作为返回值标识
    ParentFrm.Hide   ' 2. 隐藏窗体（注意：不是 Unload），这样 getUIres 才能继续读取其他控件的值
End Sub

Function getUIres() 'imdl
     Dim DecCnt, DecCode
     Dim Apc As Object: Set Apc = KCL.GetApc()
         Dim ExecPjt As Object: Set ExecPjt = Apc.ExecutingProject
         On Error Resume Next
          Dim mdl: Set mdl = ExecPjt.VBProject.VBE.Activecodepane.codemodule
             Error.Clear
         On Error GoTo 0
    If mdl Is Nothing Then Exit Function
        DecCnt = mdl.CountOfDeclarationLines ' 获取声明行数
    If DecCnt < 1 Then Exit Function
        DecCode = mdl.Lines(1, DecCnt) ' 获取声明代码
    
    Dim clscfg: Set clscfg = ParseCts(DecCode)
    Dim ttl: ttl = ParseTitle(DecCode)
    Load WD
    WD.Tag = ""
    Call WD.setFrm(ttl, clscfg)
' --- 3. 绑定按钮事件 ---
    Dim eventCol As New collection ' 必须保持这个集合存活，否则事件对象会被销毁
    Dim ctl As Control
    Dim evtHandler
    For Each ctl In WD.controls
        If TypeName(ctl) = "CommandButton" Then
            Set evtHandler = New Cls_DynaFrm
            Set evtHandler.ControlBtn = ctl
            Set evtHandler.ParentFrm = WD
            evtHandler.ControlID = ctl.Name
            eventCol.Add evtHandler
        End If
    Next
      ' --- 4. 显示窗体 (模态) ---
    WD.Show vbModal     ' 代码会在这里暂停，直到窗体被 Hide
     ' --- 5. 收集返回值 ---
    Dim Res As Object
    Set Res = CreateObject("Scripting.Dictionary")
    ' 记录点击了哪个按钮 (通过 Tag 传递)
    Res.Add "btn_clicked", WD.Tag
    
    
    
    ' 遍历所有控件获取值
    For Each ctl In WD.controls
        On Error Resume Next ' 防止某些控件没有 Value 或 Text 属性
        If TypeName(ctl) = "CheckBox" Then
            Res.Add ctl.Name, ctl.value
        ElseIf TypeName(ctl) = "TextBox" Then
            Res.Add ctl.Name, ctl.Text
        End If
        On Error GoTo 0
    Next
    ' --- 6. 清理 ---
    Unload WD
    Set getUIres = Res
End Function

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
    Dim lst: Set lst = InitLst
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

