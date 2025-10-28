Attribute VB_Name = "ASM_deleteChildren"
'Attribute VB_Name = "M37_DeleteChildren"
' ����
'{GP:3}
'{EP:DeleteChildren}
'{Caption:ɾ���Ӳ�Ʒ}
'{ControlTipText: һ��ɾ��ѡ��Ĳ�Ʒ���Ӳ�Ʒ}
'{BackColor:}
' ����ģ�鼶����

Sub DeleteChildren()

  If CATIA.Windows.Count < 1 Then
        MsgBox "û�д򿪵Ĵ���"
        Exit Sub
    End If
  If Not CanExecute("ProductDocument") Then Exit Sub
 
  Dim btn, imsg, bTitle, bResult
   imsg = "ѡ�񸸼���ɾ���������Ӳ�Ʒ�������ʹ��,�Ƿ����"
  btn = vbYesNo + vbExclamation
  bResult = MsgBox(imsg, btn, "bTitle")
     
        Select Case bResult ' Yes(6),No(7),cancel(2)
          
            Case 7 '===ѡ�񡰷�====
                Exit Sub
            Case 6  '===ѡ���ǡ�,���в�Ʒѡ��====
              Dim filter(0), iSel
                Set oDoc = CATIA.ActiveDocument
                Set osel = CATIA.ActiveDocument.Selection
            
                imsg = "��ѡ�񸸼�"
                filter(0) = "Product"
                Set iSel = KCL.SelectElement(imsg, filter).Value
                If iSel Is Nothing Then Exit Sub
                
            For Each Prd In iSel.Products
              osel.Add Prd
            Next
          
             imsg = "��ɾ��" & iSel.PartNumber & iSel.Name & "�������Ӳ�Ʒ����ȷ����"
             
             bResult = MsgBox(imsg, btn, "bTitle")
             Select Case bResult
                Case 7 '===ѡ�񡰷�====
                    Exit Sub
                Case 6  '===ѡ���ǡ�,���в�Ʒѡ��====
                  
            On Error Resume Next
                    osel.Delete
                    osel.Clear
           On Error GoTo 0
            End Select
            
        End Select

    
    
End Sub
   

