Attribute VB_Name = "m42_initme"
'Attribute VB_Name = "initme"
'{GP:4}
'{Ep:initme}
'{Caption:��ʼ��ģ��}
'{ControlTipText:��ѡ��Ĳ�Ʒ���Ӳ�Ʒ�ĵ���ģ���ʽ��}
'{BackColor:16744703}

Sub initme()

Set pdm = New class_PDM
    Set allPN = KCL.InitDic(vbTextCompare)
    allPN.RemoveAll
    
            Dim oPrd
        
'    If Not gPrd Is Nothing Then

        Set oPrd = rootPrd
        If allPN.Exists(oPrd.PartNumber) = False Then
            allPN(oPrd.PartNumber) = 1
            Call pdm.initPrd(oPrd)
        End If
            For Each product In oPrd.Products
                Call pdm.initPrd(product)
        Next
            allPN.RemoveAll
'    Else
'            MsgBox "����ѡ��Ҫ��ʼ���Ĳ�Ʒ"
'    End If
End Sub
