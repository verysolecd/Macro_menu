Attribute VB_Name = "OTH_PrePn"
'{GP:6}
'{Ep:Pnmgr}
'{Caption:零件号管理}
'{ControlTipText:零件号批量管理}
'{BackColor:}
'控件格式为 %UI <ControlType> <ControlName> <Caption/Text>
' %UI Label lbL_jpzcs  键盘造车手出品
' %UI TextBox  txt_oldstr 输入被替换字符串或留空
' %UI TextBox  txt_str 字符串
' %UI CheckBox chk_prefix  字符串增加为前缀
' %UI CheckBox  chk_suffix  字符串增加为后缀
' %UI CheckBox  chk_rep  替换零件号内字符串
' %UI CheckBox chk_delete  删除零件号内字符串
' %UI Button btnOK  确定
' %UI Button btncancel  取消
' 强制变量声明（核心修复）
Option Explicit

Private prj, mWD
Private allPN As Object
Private Const mdlname As String = "OTH_PrePn"

Sub Pnmgr()
    If Not KCL.CanExecute("ProductDocument,partdocument") Then Exit Sub
    If pdm Is Nothing Then Set pdm = New Cls_PDM
    If pdm.CurrentProduct Is Nothing Then Set pdm.CurrentProduct = KCL.defPrd
    Dim oprd As Object: Set oprd = pdm.CurrentProduct
    If oprd Is Nothing Then Exit Sub
    Set mWD = New Cls_DynaWD   '创建窗口实例
    mWD.getUIcfgfromModDEC mdlname '获取窗口控件配置
    mWD.Show '显示窗口
    Dim istr As String, oldstr As String
    Select Case mWD.btnClicked
        Case "btnOK"
            istr = ""
            If mWD.Results("txt_str") <> "" And Not KCL.ExistsKey(mWD.Results("txt_str"), "字符") Then
                istr = mWD.Results("txt_str")  'Trim(mWD.Results("txt_str"))
            End If
            If mWD.Results("txt_oldstr") <> "" And Not KCL.ExistsKey(mWD.Results("txt_oldstr"), "字符") Then
                oldstr = mWD.Results("txt_oldstr")  'Trim(mWD.Results("txt_str"))
            End If
 
            If istr = "" Then
                MsgBox "请输入有效字符串！", vbExclamation
                Exit Sub
            End If
            Set allPN = KCL.InitDic
            If mWD.Results("chk_prefix") Then
                Call c_pn_Prefix(oprd, istr)
            ElseIf mWD.Results("chk_suffix") Then
            istr = "_Rev" & istr & "_"
                Call c_pn_suffix(oprd, istr)
            ElseIf mWD.Results("chk_delete") Then
                Call del_pn_midx(oprd, istr)
            ElseIf mWD.Results("chk_rep") Then
                Call rep_pn_midx(oprd, oldstr, istr)
            End If
            Set allPN = Nothing
            MsgBox "零件号批量修改完成！", vbInformation
        Case Else: Exit Sub
    End Select
End Sub

Private Sub c_pn_Prefix(oprd As Object, istr As String)
    Dim pn As String, purePN As String, newPn As String
    Dim childProduct As Object
    pn = oprd.ReferenceProduct.partNumber
    If Not allPN.Exists(pn) Then
        allPN(pn) = 1
        purePN = KCL.StrAF(pn, "_._")
        newPn = istr & "_._" & purePN
        oprd.partNumber = newPn
        allPN(newPn) = 1 ' 记录新零件号
    End If
    If oprd.Products.count > 0 Then
        For Each childProduct In oprd.Products
            Call c_pn_Prefix(childProduct, istr)
        Next
    End If
End Sub

Private Sub c_pn_suffix(oprd As Object, istr As String)
    Dim pn As String, newPn As String, purePN As String
    Dim childProduct As Object
    pn = oprd.ReferenceProduct.partNumber
    If Not allPN.Exists(pn) Then
        allPN(pn) = 1
        purePN = KCL.StrBF(pn, "_._")
        newPn = purePN & "_._" & istr
        oprd.partNumber = newPn
        allPN(newPn) = 1
    End If
    If oprd.Products.count > 0 Then
        For Each childProduct In oprd.Products
            Call c_pn_suffix(childProduct, istr)
        Next
    End If
End Sub
Private Sub del_pn_midx(oprd As Object, istr As String)
    Dim pn As String, newPn As String
    Dim childProduct As Object
    pn = oprd.ReferenceProduct.partNumber
    If Not allPN.Exists(pn) Then
        allPN(pn) = 1
        newPn = Replace(pn, istr, "")
        If newPn <> "" Then
            oprd.partNumber = newPn
        End If
        allPN(newPn) = 1
    End If
    If oprd.Products.count > 0 Then
        For Each childProduct In oprd.Products
            Call del_pn_midx(childProduct, istr)
        Next
    End If
End Sub

Private Sub rep_pn_midx(oprd, oldstr, istr)
    Dim pn As String, newPn As String
    Dim childProduct As Object
    pn = oprd.ReferenceProduct.partNumber
    If Not allPN.Exists(pn) Then
        allPN(pn) = 1
        newPn = Replace(pn, oldstr, istr)
        If newPn <> "" Then
            oprd.partNumber = newPn
        End If
        allPN(newPn) = 1
    End If
    If oprd.Products.count > 0 Then
        For Each childProduct In oprd.Products
            Call rep_pn_midx(childProduct, oldstr, istr)
        Next
    End If
End Sub
