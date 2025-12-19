Attribute VB_Name = "A00_mdl2wd"


Option Explicit

Function mdl2wd(imdl)

Dim mdl, DecCnt, DecCode

Set mdl = imdl
    If mdl Is Nothing Then Exit Function
    DecCnt = mdl.CountOfDeclarationLines ' 获取声明行数
        If DecCnt < 1 Then Exit Function
        DecCode = mdl.Lines(1, DecCnt) ' 获取声明代码
    Dim clscfg:   Set clscfg = ParseCts(DecCode)
    Dim ttl: ttl = ParseTitle(DecCode)
'    Dim frm: Set frm = wd
    Load wd
    Call wd.setFrm(ttl, clscfg)

' --- 3. 绑定按钮事件 ---
    Dim eventCol As New collection ' 必须保持这个集合存活，否则事件对象会被销毁
    Dim ctl As Control
    Dim evtHandler As Cls_Ctrl
    For Each ctl In wd.controls
        If TypeName(ctl) = "CommandButton" Then
            Set evtHandler = New Cls_Ctrl
            Set evtHandler.ControlBtn = ctl
            Set evtHandler.ParentFrm = wd
            evtHandler.ControlID = ctl.Name
            eventCol.Add evtHandler
        End If
    Next
    
      ' --- 4. 显示窗体 (模态) ---
    ' 代码会在这里暂停，直到窗体被 Hide
    wd.Show vbModal
    
     ' --- 5. 收集返回值 ---
    Dim res As Object
    Set res = CreateObject("Scripting.Dictionary")
    ' 记录点击了哪个按钮 (通过 Tag 传递)
    res.Add "Status", wd.Tag
    ' 遍历所有控件获取值
    For Each ctl In wd.controls
        On Error Resume Next ' 防止某些控件没有 Value 或 Text 属性
        If TypeName(ctl) = "CheckBox" Then
            res.Add ctl.Name, ctl.value
        ElseIf TypeName(ctl) = "TextBox" Then
            res.Add ctl.Name, ctl.Text
        End If
        On Error GoTo 0
    Next
    ' --- 6. 清理 ---
    Unload wd
    Set mdl2wd = res
    
    
'    resultAry = wdCfg()
End Function
Function getmdlname()
   getmdlname = ""
    Dim Apc As Object: Set Apc = KCL.GetApc()
    Dim ExecPjt As Object: Set ExecPjt = Apc.ExecutingProject
    On Error Resume Next
     Set mdl = ExecPjt.VBProject.VBE.Activecodepane.codemodule
        Error.Clear
    On Error GoTo 0
    If mdl Is Nothing Then Exit Function
    getmdlname = mdl.Name
End Function
Function getiType(ctrltypename)
    Select Case LCase(ctrltypename)
        Case "commandbutton", "button", "cmd", "cbt"
            getiType = "Forms.CommandButton.1"
        Case "textbox", "text", "txt", "txt"
            getiType = "Forms.TextBox.1"
        Case "label", "lbl"
            getiType = "Forms.Label.1"
        Case "checkbox", "check", "chk"
            getiType = "Forms.CheckBox.1"
        Case "optionbutton", "option", "opt"
            getiType = "Forms.OptionButton.1"
        Case "listbox", "list", "lst"
            getiType = "Forms.ListBox.1"
        Case "combobox", "combo", "cmb"
            getiType = "Forms.ComboBox.1"
        Case "multipage", "multipages", "mpg"
            getiType = "Forms.MultiPage.1"
        Case Else
            ' 默认返回文本框类型
            getiType = "Forms.TextBox.1"
    End Select
End Function
Private Function ParseCts(ByVal code As String) As Object
    Dim regex As Object
    Dim matches As Object
    Dim match As Object
    Dim Cls_property As Object ' Scripting.Dictionary
    Set regex = CreateObject("VBScript.RegExp")
    With regex
        .Global = True
        .MultiLine = True
        ' 格式: ' %UI <ControlType> <ControlName> <Caption/Text>
        .Pattern = "^\s*'\s*%UI\s+(\w+)\s+(\w+)\s+(.*)$"
    End With
    
    Dim lst
    Set lst = InitLst
    If regex.TEST(code) Then
        Set matches = regex.Execute(code)
        For Each match In matches
            Set Cls_property = KCL.InitDic
            Cls_property.Add "Type", getiType(match.SubMatches(0))
            Cls_property.Add "Name", match.SubMatches(1)
            Cls_property.Add "Caption", Trim(match.SubMatches(2))
'            If Not ctrlcoll.Contains(Cls_property("Name")) Then
               lst.Add Cls_property '("Name") , Cls_property
'            End If
        Next
    End If
    Set ParseCts = lst
End Function


Private Function ParseTitle(ByVal code As String) As String
        Dim regex As Object
        Dim matches As Object
        Dim match As Object
        Dim itl
        Set regex = KCL.getRegexp
    With regex
        .Global = True
        .MultiLine = True
        ' 格式: ' %UI <Title> <Caption/Text>
        .Pattern = "^\s*'\s*%Title\s+(.*)$"
    End With
    
    itl = "请问你要如何执行配置"
              If regex.TEST(code) Then
                Set matches = regex.Execute(code)
                For Each match In matches
                 itl = match.SubMatches(0)
                Next
              End If
    ParseTitle = itl
End Function








