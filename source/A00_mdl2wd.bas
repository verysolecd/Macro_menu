Attribute VB_Name = "A00_mdl2wd"


Option Explicit

Sub mdl2wd(imdl)

Dim mdl, DecCnt, DecCode
'regPtn = TAG_S & "(.*?)" & TAG_D & "(.*?)" & TAG_E
'Set inf = KCL.getInfo_asDic(Ctrlinf, regPtn)
'    Dim Apc As Object: Set Apc = KCL.GetApc()
'    Dim ExecPjt As Object: Set ExecPjt = Apc.ExecutingProject
'    On Error Resume Next
'     Set mdl = ExecPjt.VBProject.VBE.Activecodepane.codemodule
'        Error.Clear
'    On Error GoTo 0
Set mdl = imdl
    If mdl Is Nothing Then Exit Sub
    DecCnt = mdl.CountOfDeclarationLines ' 获取声明行数
        If DecCnt < 1 Then Exit Sub
        DecCode = mdl.Lines(1, DecCnt) ' 获取声明代码
    Dim clscfg:   Set clscfg = ParseCts(DecCode)
    Dim ttl: ttl = ParseTitle(DecCode)
    Dim frm: Set frm = wd
    Call frm.setFrm(ttl, clscfg)
    frm.Show vbModeless
'    resultAry = wdCfg()
End Sub
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








