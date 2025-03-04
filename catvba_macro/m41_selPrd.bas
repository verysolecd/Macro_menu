Attribute VB_Name = "m41_selPrd"
'Attribute VB_Name = "selPrd"
'{GP:4}
'{Ep:whois2rv}
'{Caption:选择产品}
'{ControlTipText:选择要被读取或修改的产品}
'{BackColor:16744703}

'
'Sub selPrd()
'
'
'Call selPrd
'
'End Sub



Function selPrd()
   Dim oSel, status, iType(0)
            MsgBox "请选择要读取的产品"
            xlApp.Windows(1).WindowState = xlMinimized
            CATIA.ActiveWindow.WindowState = 0
            Set oSel = CATIA.ActiveDocument.Selection
            oSel.Clear
            iType(0) = "Product"
            If oSel.Count2 = 0 Then
                status = oSel.SelectElement2(iType, "请选择要读取的产品", False)
                'status = oSel.SelectElement3(iType, "请选择要读取的产品", True, 2, False)
            End If
            If status = "Cancel" Then
                xlApp.Windows(1).WindowState = xlMaximized
                Exit Function
            End If
            If status = "Normal" And oSel.Count2 = 1 Then
                    Set selPrd = oSel.Item(1).LeafProduct.ReferenceProduct
                    oSel.Clear
            Else
                MsgBox "请只选择一个产品"
                Exit Function
            End If
End Function
