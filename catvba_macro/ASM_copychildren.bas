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
If CATIA.Windows.Count < 1 Then
    MsgBox "没有打开的窗口"
    Exit Sub
End If
If Not CanExecute("ProductDocument") Then Exit Sub

Dim imsg, filter(0), iSel
Set oDoc = CATIA.ActiveDocument
Set osel = CATIA.ActiveDocument.Selection

On Error Resume Next
    imsg = "请选择要复制的子产品父集"
    MsgBox imsg
    filter(0) = "Product"
    Dim sourcePrd, targetPrd
    Set sourcePrd = KCL.SelectElement(imsg, filter).Value
    If sourcePrd Is Nothing Then Exit Sub
    For Each Prd In sourcePrd.Products
       osel.Add Prd
    Next
        osel.Copy
        osel.Clear
    imsg = "请选择黏贴目标产品"
    MsgBox imsg
    Set targetPrd = KCL.SelectElement(imsg, filter).Value
    If targetPrd Is Nothing Then
        Exit Sub
    Else
        osel.Add targetPrd
        osel.Paste
        Set targetPrd = Nothing
    End If
On Error GoTo 0
End Sub

