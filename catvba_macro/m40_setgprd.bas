Attribute VB_Name = "m40_setgprd"
'{GP:4}
'{Ep:setgprd}
'{Caption:选择产品}
'{ControlTipText:选择要被读取或修改的产品}
'{BackColor:16744703}


Sub setgprd()
'            Dim oSel, status, iType(0)
'            CATIA.ActiveWindow.WindowState = 0
'            MsgBox "请选择要读取的产品"
'
'            Set oSel = CATIA.Activedocument.Selection
'            oSel.Clear
'            iType(0) = "Product"
'            If oSel.Count2 = 0 Then
'                status = oSel.SelectElement2(iType, "请选择要读取的产品", False)
'                status = oSel.SelectElement3(iType, "请选择要读取的产品", True, 2, False)
'            End If
'            If status = "Cancel" Then
'                Exit Function
'            End If
'            If status = "Normal" And oSel.Count2 = 1 Then
'                    Set selPrd = oSel.Item(1).LeafProduct.ReferenceProduct
'                    oSel.Clear
'            Else
'                MsgBox "请选择且仅选择一个产品"
'                Exit Function
'                oSel.Clear
'            End If

    Dim xlm, pdm
    Set pdm = New class_PDM
    On Error Resume Next
     Set gPrd = pdm.defgprd()
        If gPrd Is Nothing Then
        MsgBox "已退出，程序将结束"
        Exit Sub
        Else
        MsgBox "待操作产品已经选择：" & gPrd.PartNumber
        End If
End Sub
