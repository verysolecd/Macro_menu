Attribute VB_Name = "ASM_copychildren"
'Attribute VB_Name = "M36_copychildren"
' 复制
'{GP:3}
'{EP:cpChildren}
'{Caption:复制子产品}
'{ControlTipText: 一键复制第一个产品的子产品到第二个产品子级}
'{BackColor:}
' 定义模块级变量

Sub cpChildren()

If Not CanExecute("ProductDocument") Then Exit Sub

Call KCL.setASM(False)

Dim imsg, filter(0), oSel
Set oDoc = CATIA.ActiveDocument
Set oSel = CATIA.ActiveDocument.Selection

oSel.Clear
On Error GoTo errorhandler
    imsg = "请先点击选择源父产品，再点击选择目标父产品"
    MsgBox imsg
    filter(0) = "Product"
    Dim sourcePrd, targetPrd
    Set sourcePrd = KCL.SelectItem(imsg, filter)
    If sourcePrd Is Nothing Then GoTo errorhandler
        For Each prd In sourcePrd.Products
           oSel.Add prd
        Next
    oSel.Copy
    oSel.Clear
    imsg = "请点击选择目标父产品"
    Set targetPrd = KCL.SelectItem(imsg, filter)
    If targetPrd Is Nothing Then
      GoTo errorhandler
    Else
        oSel.Add targetPrd
        oSel.Paste
        
    End If
        oSel.Clear
        Set targetPrd = Nothing
         Set sourcePrd = Nothing
    On Error GoTo 0
    
    Call KCL.setASM(True)
    
    
errorhandler:
        If Err.Number <> 0 Then
            Call KCL.setASM(True)
              oSel.Clear
            Err.Clear
            MsgBox "CATIA 程序错误：" & Err.Description, vbCritical
        Exit Sub
        Else
        
         Call KCL.setASM(True)
        End If

End Sub

