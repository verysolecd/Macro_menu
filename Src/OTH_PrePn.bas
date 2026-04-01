Attribute VB_Name = "OTH_PrePn"
'{GP:6}
'{Ep:Pnmgr}
'{Caption:零件号管理}
'{ControlTipText:零件号批量管理}
'{BackColor:}
'控件格式为 %UI <ControlType> <ControlName> <Caption/Text>
' %UEI Label lbL_jpzcs  键盘造车手出品
' %UI TextBox  txt_str 字符串
' %UI CheckBox chk_prefix  字符串增加为前缀
' %UI CheckBox  chk_suffix  字符串增加为后缀
' %UI CheckBox chk_delete  删除零件号内字符串
' %UI Button btnOK  确定
' %UI Button btncancel  取消

' 强制变量声明（核心修复）
Option Explicit

Private prj
Private allPN As Object
Private Const mdlname As String = "OTH_PrePn"

Sub Pnmgr()
    If Not KCL.CanExecute("ProductDocument,partdocument") Then Exit Sub
    Dim oPrd As Object
    Set oPrd = CATIA.ActiveDocument.Product
    If oPrd Is Nothing Then Exit Sub
    Dim oFrm As Object
    Set oFrm = KCL.newFrm(mdlname)
    oFrm.Show
    Dim istr As String
    Select Case oFrm.BtnClicked
        Case "btnOK"
            istr = ""
            If oFrm.res("txt_str") <> "" And Not KCL.ExistsKey(oFrm.res("txt_str"), "字符") Then
                istr = Trim(oFrm.res("txt_str"))
            End If
            If istr = "" Then
                MsgBox "请输入有效字符串！", vbExclamation
                Exit Sub
            End If
            Set allPN = KCL.InitDic
            If oFrm.res("chk_prefix") Then
                Call c_pn_Prefix(oPrd, istr)
            ElseIf oFrm.res("chk_suffix") Then
            istr = "_Rev" & istr & "_"
                Call c_pn_suffix(oPrd, istr)
            ElseIf oFrm.res("chk_delete") Then
                Call del_pn_midx(oPrd, istr)
            End If
            Set allPN = Nothing
            MsgBox "零件号批量修改完成！", vbInformation
        Case Else: Exit Sub
    End Select
End Sub


Private Sub c_pn_Prefix(oPrd As Object, istr As String)
    Dim pn As String, purePN As String, newPn As String
    Dim childProduct As Object
    pn = oPrd.ReferenceProduct.partNumber
    If Not allPN.Exists(pn) Then
        allPN(pn) = 1
        purePN = KCL.StrAF(pn, "_._")
        newPn = istr & "_._" & purePN
        oPrd.partNumber = newPn
        allPN(newPn) = 1 ' 记录新零件号
    End If
    If oPrd.Products.count > 0 Then
        For Each childProduct In oPrd.Products
            Call c_pn_Prefix(childProduct, istr)
        Next
    End If
End Sub

Private Sub c_pn_suffix(oPrd As Object, istr As String)
    Dim pn As String, newPn As String
    Dim childProduct As Object
    pn = oPrd.ReferenceProduct.partNumber
    If Not allPN.Exists(pn) Then
        allPN(pn) = 1
        purePN = KCL.StrBF(pn, "_._")
        newPn = purePN & "_._" & istr
        oPrd.partNumber = newPn
        allPN(newPn) = 1
    End If
    If oPrd.Products.count > 0 Then
        For Each childProduct In oPrd.Products
            Call c_pn_suffix(childProduct, istr)
        Next
    End If
End Sub
Private Sub del_pn_midx(oPrd As Object, istr As String)
    Dim pn As String, newPn As String
    Dim childProduct As Object
    pn = oPrd.ReferenceProduct.partNumber
    If Not allPN.Exists(pn) Then
        allPN(pn) = 1
        newPn = Replace(pn, istr, "")
        If newPn <> "" Then
            oPrd.partNumber = newPn
        End If
        allPN(newPn) = 1
    End If
    If oPrd.Products.count > 0 Then
        For Each childProduct In oPrd.Products
            Call del_pn_midx(childProduct, istr)
        Next
    End If
End Sub
