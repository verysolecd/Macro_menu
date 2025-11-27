Attribute VB_Name = "A000000000000_test"

'
'格式为 %UI <ControlType> <ControlName> <Caption/Text>

' %UI Label lbL_jpzcs 键盘造车手出品
' %UI CheckBox chk_path  是否导出到当前路径
' %UI CheckBox  chk_tm  是否更新时间戳到CATIA零件号？
' %UI TextBox   txt_tm  请输入时间格式
' %UI CheckBox chk_log  本次导出是否更新日志？
' %UI TextBox   txt_log  请输入更新内容(不必输入时间)
' %UI Button btnOK  确定
' %UI Button btncancel  取消

Option Explicit


Sub TEST()
Dim mdl, DecCnt, DecCode
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
 Dim inf:   Set inf = ParseCts(DecCode)
 Dim frm: Set frm = wd
 Call frm.setFrm(inf)
    frm.Show
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

