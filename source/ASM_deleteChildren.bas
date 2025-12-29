Attribute VB_Name = "ASM_deleteChildren"
'Attribute VB_Name = "M37_DeleteChildren"
' 复制
'{GP:3}
'{EP:DeleteChildren}
'{Caption:删除子产品}
'{ControlTipText: 一键删除选择的产品的子产品}
'{BackColor:}
' 定义模块级变量
Option Explicit

Sub DeleteChildren()
    If Not CanExecute("ProductDocument") Then Exit Sub
    
    Dim oSel: Set oSel = CATIA.ActiveDocument.Selection: oSel.Clear

    Dim imsg, filter(0), iSel
      imsg = "请选择父集": filter(0) = "Product"
       Set iSel = KCL.SelectItem(imsg, filter)
    If iSel Is Nothing Then Exit Sub
    Dim prd
    For Each prd In iSel.Products
      oSel.Add prd
    Next
    
      Dim btn, bTitle, bResult
      imsg = "将删除" & iSel.PartNumber & iSel.Name & "下的所有子产品，您确认吗"
      btn = vbYesNo + vbExclamation
      bResult = MsgBox(imsg, btn, "bTitle")  ' Yes(6),No(7),cancel(2)
           Select Case bResult
              Case 7: Exit Sub '===选择“否”====
              Case 6  '===选择“是”,进行产品选择====
                  On Error Resume Next
                       oSel.Delete
                       oSel.Clear
                  On Error GoTo 0
          End Select

End Sub
