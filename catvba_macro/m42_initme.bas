Attribute VB_Name = "m42_initme"
'Attribute VB_Name = "initme"
'{GP:4}
'{Ep:initme}
'{Caption:��ʼ��ģ��}
'{ControlTipText:��ѡ��Ĳ�Ʒ���Ӳ�Ʒ�ĵ���ģ���ʽ��}
'{BackColor:16744703}

Sub initme()

     If pdm Is Nothing Then
          Set pdm = New class_PDM
     End If
     
     Set allPN = KCL.InitDic(vbTextCompare)
     allPN.RemoveAll
     
    Dim oPrd
'    If Not gPrd Is Nothing Then
'        Set oPrd = gPrd
        Set oPrd = CATIA.ActiveDocument.Product
'
'    End If
    
     Call ini_oPrd(oPrd)
     
     allPN.RemoveAll
     MsgBox "���ģ���Ѿ�Ӧ��"
'
'    Else
'        MsgBox "����ѡ��Ҫ��ʼ���Ĳ�Ʒ"
'    End If
    
End Sub

Sub ini_oPrd(oPrd)

        If allPN.Exists(oPrd.PartNumber) = False Then
            allPN(oPrd.PartNumber) = 1
            Call pdm.initPrd(oPrd)
        End If
            For Each Product In oPrd.Products
                Call ini_oPrd(Product)
          Next
End Sub

