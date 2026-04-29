' 控件类型	ObjectType 字符串
' 命令按钮	"Forms.CommandButton.1"  CBT 
' 文本框	"Forms.TextBox.1"               txt_log
' 标签	"Forms.Label.1"  lbl
' 复选框	"Forms.CheckBox.1"  chk
' 单选按钮	"Forms.OptionButton.1"  opt
' 列表框	"Forms.ListBox.1"  lst
' 组合框	"Forms.ComboBox.1"  cmb
' 多页控件 "Forms.multipages.1"  mpg


private thistop as long
private const itemgap = 20
private const itemheight = 20
private const itemwidth  = 100


Sub TestCustomPopup()
    ' 设置窗体基本属性
    Me.Caption = "零件号更新和路径选择"
    Me.Width = 280
    Me.Height = 200
    Me.StartUpPosition = 1
    Call getctrls
    call initfrm
End Sub


sub initfrm

call setFrm(getctrlinfos(mdl))

end sub

function getctrlinfos(mdl as object)
    desc= getdesc(mdl)
    ctrlinfos= getctrtlsinfo(desc)
    getctrlinfos= ctrlinfos
end function



 call setFrm(ctrlinfos)


 for each d In  ctrlinfos
 call addCtrl(d.ctrltype, d.ctrlname, d.ctrlenabled)
   
   
next d







function addCtrl(ctrltype, ctrlname, ctrlenabled)
dim it
if isstring(ctrltype) Then
    itype = getCtrlObjectType(ctrltype)
end if

 Dim newctrl
 Set newctrl = Me.controls.Add(itype, ctrlname, ctrlenabled)
 Set addCtrl = newctrl
 
 with new ctrl
 
    .name = ctrlname
    .enabled = ctrlenabled
    .caption = ctrlname
    .top= this.top + 30
    .heigt = 20
    .left = 20
    .width = 100

     thistop = .top + .Height + itemgap
end with

End function




End Sub

Function getCtrlObjectType(ctrltypename)
    Select Case LCase(ctrltypename)
        Case "commandbutton", "button", "cmd", "cbt"
            getCtrlObjectType = "Forms.CommandButton.1"
        Case "textbox", "text", "txt", "txt_log"
            getCtrlObjectType = "Forms.TextBox.1"
        Case "label", "lbl"
            getCtrlObjectType = "Forms.Label.1"
        Case "checkbox", "check", "chk"
            getCtrlObjectType = "Forms.CheckBox.1"
        Case "optionbutton", "option", "opt"
            getCtrlObjectType = "Forms.OptionButton.1"
        Case "listbox", "list", "lst"
            getCtrlObjectType = "Forms.ListBox.1"
        Case "combobox", "combo", "cmb"
            getCtrlObjectType = "Forms.ComboBox.1"
        Case "multipage", "multipages", "mpg"
            getCtrlObjectType = "Forms.MultiPage.1"
        Case Else
            ' 默认返回文本框类型
            getCtrlObjectType = "Forms.TextBox.1"
    End Select
End Function





AI信息


Sub TestCustomPopup()
    Dim ctrlItems
    Dim i
    
    ' 解析当前模块中的控件声明信息
    ctrlItems = ParseModuleDeclarations()
    
    ' 根据解析结果创建控件
    For i = 0 To UBound(ctrlItems) Step 3
        Dim ctrlType, ctrlName, ctrlVisible
        ctrlType = ctrlItems(i)
        ctrlName = ctrlItems(i + 1)
        ctrlVisible = ctrlItems(i + 2)
        
        ' 调用添加控件函数
        Call addCtrl(ctrlType, ctrlName, ctrlVisible)
    Next i
    
    ' 设置控件属性和布局
    SetupControlsLayout
End Sub

Function ParseModuleDeclarations()
    ' 这里应该从当前模块读取声明信息
    ' 示例数据结构，实际应该从模块代码中解析
    Dim ctrlItems
    ctrlItems = Array( _
        "textbox", "txt_log", True, _
        "commandbutton", "cmd_ok", True, _
        "label", "lbl_title", True, _
        "combobox", "cmb_options", True _
    )
    ParseModuleDeclarations = ctrlItems
End Function

Function addCtrl(ctrltype, ctrlname, ctrlenabled)
    Dim newctrl
    
    If IsString(ctrltype) Then
        ' 根据控件类型字符串确定 ObjectType
        Set newctrl = Me.Controls.Add(getCtrlObjectType(ctrltype), ctrlname, ctrlenabled)
    Else
        ' 如果 ctrltype 直接就是 ObjectType 字符串
        Set newctrl = Me.Controls.Add(ctrltype, ctrlname, ctrlenabled)
    End If
    
    Set addCtrl = newctrl
End Function

' 辅助函数：根据控件类型名称返回对应的 ObjectType 字符串
Function getCtrlObjectType(ctrltypename)
    Select Case LCase(ctrltypename)
        Case "commandbutton", "button", "cmd", "cbt"
            getCtrlObjectType = "Forms.CommandButton.1"
        Case "textbox", "text", "txt", "txt_log"
            getCtrlObjectType = "Forms.TextBox.1"
        Case "label", "lbl"
            getCtrlObjectType = "Forms.Label.1"
        Case "checkbox", "check", "chk"
            getCtrlObjectType = "Forms.CheckBox.1"
        Case "optionbutton", "option", "opt"
            getCtrlObjectType = "Forms.OptionButton.1"
        Case "listbox", "list", "lst"
            getCtrlObjectType = "Forms.ListBox.1"
        Case "combobox", "combo", "cmb"
            getCtrlObjectType = "Forms.ComboBox.1"
        Case "multipage", "multipages", "mpg"
            getCtrlObjectType = "Forms.MultiPage.1"
        Case Else
            ' 默认返回文本框类型
            getCtrlObjectType = "Forms.TextBox.1"
    End Select
End Function

Sub SetupControlsLayout()
    ' 设置控件的基本布局和属性
    Dim ctrl As Object
    
    ' 设置日志文本框
    If Me.Controls.Exists("txt_log") Then
        Set ctrl = Me.txt_log
        ctrl.Top = 30
        ctrl.Left = 10
        ctrl.Width = 200
        ctrl.Height = 100
        ctrl.MultiLine = True
        ctrl.ScrollBars = fmScrollBarsVertical
    End If
    
    ' 设置确定按钮
    If Me.Controls.Exists("cmd_ok") Then
        Set ctrl = Me.cmd_ok
        ctrl.Top = 150
        ctrl.Left = 70
        ctrl.Width = 80
        ctrl.Caption = "确定"
    End If
    
    ' 设置标题标签
    If Me.Controls.Exists("lbl_title") Then
        Set ctrl = Me.lbl_title
        ctrl.Top = 10
        ctrl.Left = 10
        ctrl.Width = 150
        ctrl.Font.Bold = True
        ctrl.Caption = "自定义弹窗"
    End If
    
    ' 设置组合框
    If Me.Controls.Exists("cmb_options") Then
        Set ctrl = Me.cmb_options
        ctrl.Top = 140
        ctrl.Left = 10
        ctrl.Width = 150
        ctrl.AddItem "选项1"
        ctrl.AddItem "选项2"
        ctrl.AddItem "选项3"
    End If
End Sub

' 控件事件处理示例
Private Sub cmd_ok_Click()
    MsgBox "确定按钮被点击"
    Me.Hide
End Sub

' 辅助函数：检查变量是否为字符串类型
Function IsString(var As Variant) As Boolean
    IsString = (VarType(var) = vbString)
End Function