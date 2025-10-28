Attribute VB_Name = "ASM_deleteChildren"
'Attribute VB_Name = "M37_DeleteChildren"
' 复制
'{GP:3}
'{EP:DeleteChildren}
'{Caption:删除子产品}
'{ControlTipText: 一键删除选择的产品的子产品}
'{BackColor:}
' 定义模块级变量

Sub DeleteChildren()

  If CATIA.Windows.Count < 1 Then
        MsgBox "没有打开的窗口"
        Exit Sub
    End If
  If Not CanExecute("ProductDocument") Then Exit Sub
 
  Dim btn, imsg, bTitle, bResult
   imsg = "选择父集后将删除其所有子产品，请谨慎使用,是否继续"
  btn = vbYesNo + vbExclamation
  bResult = MsgBox(imsg, btn, "bTitle")
     
        Select Case bResult ' Yes(6),No(7),cancel(2)
          
            Case 7 '===选择“否”====
                Exit Sub
            Case 6  '===选择“是”,进行产品选择====
              Dim filter(0), iSel
                Set oDoc = CATIA.ActiveDocument
                Set osel = CATIA.ActiveDocument.Selection
            
                imsg = "请选择父集"
                filter(0) = "Product"
                Set iSel = KCL.SelectElement(imsg, filter).Value
                If iSel Is Nothing Then Exit Sub
                
            For Each Prd In iSel.Products
              osel.Add Prd
            Next
          
             imsg = "将删除" & iSel.PartNumber & iSel.Name & "的所有子产品，您确认吗"
             
             bResult = MsgBox(imsg, btn, "bTitle")
             Select Case bResult
                Case 7 '===选择“否”====
                    Exit Sub
                Case 6  '===选择“是”,进行产品选择====
                  
            On Error Resume Next
                    osel.Delete
                    osel.Clear
           On Error GoTo 0
            End Select
            
        End Select

    
    
End Sub
   

