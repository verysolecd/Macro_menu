Attribute VB_Name = "A000000000000_test"
' 控件类型  ObjectType 字符串
' 命令按钮  "Forms.CommandButton.1"  CBT
' 文本框    "Forms.TextBox.1"               txt_log
' 标签  "Forms.Label.1"  lbl
' 复选框    "Forms.CheckBox.1"  chk
' 单选按钮  "Forms.OptionButton.1"  opt
' 列表框    "Forms.ListBox.1"  lst
' 组合框    "Forms.ComboBox.1"  cmb
' 多页控件 "Forms.multipages.1"  mpg
' 控件映射表
'
'%UI txt <ControlName> <Left> <Top> <Width> <Height> <Caption/Text>
' %UI Label lblName 10 40 80 20 名称:
' %UI TextBox txtName 90 38 180 22 请输入...
' %UI CheckBox chkEnable 10 70 150 20 启用高级选项
' %UI Button btnOK 110 110 80 25 确定
' %UI Button btncancel 200 110 80 25 取消
Private TagMap As Object                    ' 分组编号标签
Private Const itemgap = 6
Private bttop
Private Const TAG_S = "{"                   ' 控件开始标签
Private Const TAG_D = ":"                   ' 控件分隔标签
Private Const TAG_E = "}"                   ' 控件结束标签
Private Const TAG_indx = "gp"              ' 分组编号标签
Private Const TAG_ENTRYPNT = "ep"           ' 入口点标签
Private Const TAG_ENTRY_DEF = "CATMain"     ' 入口点默认值
Private Const TAG_PJTPATH = "pjt_path"      ' 项目路径标签
Private Const TAG_MDLNAME = "mdl_name"      ' 模块名称标签
Private Const Ctrlinf = _
            "{frm: 导出stp选项 }" & _
            "{txt_log: txt }" & _
            "{chk_tm : chk }" & _
            "{chk_pn : chk }"
Sub TEST()
'regPtn = TAG_S & "(.*?)" & TAG_D & "(.*?)" & TAG_E
'Set inf = KCL.getInfo_asDic(Ctrlinf, regPtn)
 Dim Apc As Object: Set Apc = KCL.GetApc()
    Dim ExecPjt As Object: Set ExecPjt = Apc.ExecutingProject
    On Error Resume Next
     Set mdl = ExecPjt.VBProject.VBE.Activecodepane.codemodule
        Error.Clear
    On Error GoTo 0
    If mdl Is Nothing Then Exit Sub
    DecCnt = mdl.CountOfDeclarationLines ' 获取声明行数
        If DecCnt < 1 Then Exit Sub
        DecCode = mdl.Lines(1, DecCnt) ' 获取声明代码
Set inf = ParseCts(DecCode)
showdict inf
thistop = 5
    Set frm = wd
     With wd
        .Caption = "我请问呢"
        .Width = 400
'        .BackColor =
        .Font.Name = Thoma
        .Font.Size = 12
        .StartUpPosition = 2 ' 居中
    End With
     Dim keys
    keys = inf.keys
    btttop = 9
    For i = 0 To UBound(keys)
         cname = keys(i)
         ctype = inf(keys(i))("Type")
        Set ctr = frm.controls.Add(ctype, cname, True)
         With ctr
             If ctype <> "Forms.TextBox.1" Then
                .Caption = cname
             End If
                 .top = thistop
                If ctype = "Forms.CommandButton.1" Then
                      If btttop = 9 Then
                          bttop = .top
                        Else
                        .top = bttop
                        End If
                End If
                .Height = 25
                .Left = 20
                .Width = 90
                 thistop = .top + .Height + 6
            End With
        Next
  wd.Show
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
Function getcType(ctrltypename)
    Select Case LCase(ctrltypename)
        Case "commandbutton", "button", "cmd", "cbt"
            getcType = "Forms.CommandButton.1"
        Case "textbox", "text", "txt", "txt"
            getcType = "Forms.TextBox.1"
        Case "label", "lbl"
            getcType = "Forms.Label.1"
        Case "checkbox", "check", "chk"
            getcType = "Forms.CheckBox.1"
        Case "optionbutton", "option", "opt"
            getcType = "Forms.OptionButton.1"
        Case "listbox", "list", "lst"
            getcType = "Forms.ListBox.1"
        Case "combobox", "combo", "cmb"
            getcType = "Forms.ComboBox.1"
        Case "multipage", "multipages", "mpg"
            getcType = "Forms.MultiPage.1"
        Case Else
            ' 默认返回文本框类型
            getcType = "Forms.TextBox.1"
    End Select
End Function
Private Function ParseCts(ByVal code As String) As Object
    Dim regex As Object
    Dim matches As Object
    Dim match As Object
    Dim Cprop As Object ' Scripting.Dictionary
    Set regex = CreateObject("VBScript.RegExp")
    With regex
        .Global = True
        .MultiLine = True
        ' 格式: ' %UI <ControlType> <ControlName> <Left> <Top> <Width> <Height> <Caption/Text>
        .Pattern = "^\s*'\s*%UI\s+(\w+)\s+(\w+)\s+([0-9\.]+)\s+([0-9\.]+)\s+([0-9\.]+)\s+([0-9\.]+)\s+(.*)$"
    End With
    Set ctrlcoll = KCL.InitDic
    If regex.TEST(code) Then
        Set matches = regex.Execute(code)
        For Each match In matches
            Set Cprop = KCL.InitDic
            Cprop.Add "Type", getcType(match.SubMatches(0))
            Cprop.Add "Name", match.SubMatches(1)
            Cprop.Add "Left", CDbl(match.SubMatches(2))
            Cprop.Add "Top", CDbl(match.SubMatches(3))
            Cprop.Add "Width", CDbl(match.SubMatches(4))
            Cprop.Add "Height", CDbl(match.SubMatches(5))
            Cprop.Add "Caption", Trim(match.SubMatches(6))
            If Not ctrlcoll.Exists(Cprop("Name")) Then
                ctrlcoll.Add Cprop("Name"), Cprop
            End If
        Next
    End If
    Set ParseCts = ctrlcoll
End Function

