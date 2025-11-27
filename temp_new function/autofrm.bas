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
 ' %UI <ControlType> <ControlName> <Caption/Text>
' %UI Label lblName  标签名字
' %UI TextBox txtName  请输入...
' %UI CheckBox chkEnable  启用高级选项
' %UI Button btnOK  确定
' %UI Button btncancel  取消
' %UI CheckBox 我i  启用高级
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
        .Height = 300
    End With
     Dim keys
    keys = inf.keys
    btttop = 0
    wd.Show modeless
    For i = 0 To UBound(keys)
         iName = keys(i)
         iType = inf(keys(i))("Type")
         icaption = inf(keys(i))("Caption")
    Set ctr = frm.controls.Add(iType, iName, True)
             With ctr
            .Name = iName
                    If iType <> "Forms.TextBox.1" Then
                       .Caption = icaption
                     Else
                       .Text = icaption
                    End If
                    
                    .Left = 20
                    .Width = 120
                        
              Select Case iType
            
                    Case "Forms.CommandButton.1"
                             If bttop = 0 Then
                                  bttop = thistop
                                  .top = bttop  '98
                                  thisleft = .Left + .Width + itemgap
                                  Debug.Print .top
                             Else
                                  .top = bttop
                                  .Left = 120
                                  Debug.Print "第二按钮高度" & .top
                                 .Left = thisleft
                             End If
                             
                        Case Else
                        
                            .top = thistop
                        
                     End Select

                    .Height = 30
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
    Set ctrlcoll = KCL.InitDic
    If regex.TEST(code) Then
        Set matches = regex.Execute(code)
        For Each match In matches
            Set Cls_property = KCL.InitDic
            Cls_property.Add "Type", getiType(match.SubMatches(0))
            Cls_property.Add "Name", match.SubMatches(1)
            Cls_property.Add "Caption", Trim(match.SubMatches(2))
            If Not ctrlcoll.Exists(Cls_property("Name")) Then
                ctrlcoll.Add Cls_property("Name"), Cls_property
            End If
        Next
    End If
    Set ParseCts = ctrlcoll
End Function

Private Sub SetControlProperties(ctlTarget As Control, dictProps As Object)
    Dim strPropName As String
    Dim varPropValue As Variant    
    ' 遍历属性字典
    For Each strPropName In dictProps.Keys
        varPropValue = dictProps(strPropName)        
        ' 用 CallByName 动态设置属性（vbLet 表示赋值）
        On Error Resume Next
        CallByName ctlTarget, strPropName, vbLet, varPropValue
        On Error GoTo 0        
        ' 输出赋值结果（便于调试）
        Debug.Print "  - 赋值：属性=" & strPropName & "，值=" & varPropValue
    Next strPropName
End Sub