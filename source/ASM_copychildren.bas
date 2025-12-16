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
Dim imsg, filter(0), iSel
Set oDoc = CATIA.ActiveDocument
Set osel = CATIA.ActiveDocument.Selection
On Error Resume Next
    imsg = "请先点击选择源父产品，再点击选择目标父产品"
    MsgBox imsg
    filter(0) = "Product"
    Dim sourcePrd, targetPrd
    Set sourcePrd = KCL.SelectItem(imsg, filter)
    If sourcePrd Is Nothing Then Exit Sub
        For Each prd In sourcePrd.Products
           osel.Add prd
        Next
    osel.Copy
    osel.Clear
    Set targetPrd = KCL.SelectItem(imsg, filter)
    If targetPrd Is Nothing Then
        Exit Sub
    Else
        osel.Add targetPrd
        osel.Paste
        Set targetPrd = Nothing
    End If
On Error GoTo 0
End Sub
