Attribute VB_Name = "m41_selPrd"
'Attribute VB_Name = "selPrd"
'{GP:4}
'{Ep:whois2rv}
'{Caption:ѡ���Ʒ}
'{ControlTipText:ѡ��Ҫ����ȡ���޸ĵĲ�Ʒ}
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
            MsgBox "��ѡ��Ҫ��ȡ�Ĳ�Ʒ"
            xlApp.Windows(1).WindowState = xlMinimized
            CATIA.ActiveWindow.WindowState = 0
            Set oSel = CATIA.ActiveDocument.Selection
            oSel.Clear
            iType(0) = "Product"
            If oSel.Count2 = 0 Then
                status = oSel.SelectElement2(iType, "��ѡ��Ҫ��ȡ�Ĳ�Ʒ", False)
                'status = oSel.SelectElement3(iType, "��ѡ��Ҫ��ȡ�Ĳ�Ʒ", True, 2, False)
            End If
            If status = "Cancel" Then
                xlApp.Windows(1).WindowState = xlMaximized
                Exit Function
            End If
            If status = "Normal" And oSel.Count2 = 1 Then
                    Set selPrd = oSel.Item(1).LeafProduct.ReferenceProduct
                    oSel.Clear
            Else
                MsgBox "��ֻѡ��һ����Ʒ"
                Exit Function
            End If
End Function
