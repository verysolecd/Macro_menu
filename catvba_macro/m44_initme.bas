Attribute VB_Name = "m44_initme"
'Attribute VB_Name = "initme"
'{GP:4}
'{Ep:initme}
'{Caption:��ʼ��ģ��}
'{ControlTipText:��ѡ��Ĳ�Ʒ���Ӳ�Ʒ�ĵ���ģ���ʽ��}
'{BackColor: }

Sub initme()

     If pdm Is Nothing Then
          Set pdm = New class_PDM
     End If
     
     Set allPN = KCL.InitDic(vbTextCompare)
     allPN.RemoveAll
     
    Dim oprd
    
   If gPrd Is Nothing Then
            pdm.defgprd
    End If
    
    
    
      Set iPrd = gPrd
      
    If Not iPrd Is Nothing Then
     On Error Resume Next
      Call ini_oPrd(iPrd)
        allPN.RemoveAll
        MsgBox "���ģ���Ѿ�Ӧ��"
       
        
        If Error.Number <> 0 Then
        MsgBox "����һ������������������lisence"
        End If
        On Error GoTo 0
    End If
    
End Sub

Sub ini_oPrd(oprd)

        If allPN.Exists(oprd.PartNumber) = False Then
            allPN(oprd.PartNumber) = 1
            Call pdm.initPrd(oprd)
        End If
            For Each Product In oprd.Products
                Call ini_oPrd(Product)
        Next
End Sub


