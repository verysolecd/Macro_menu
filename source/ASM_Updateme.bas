Attribute VB_Name = "ASM_Updateme"
'Attribute VB_Name = "m33_Updateme"
'{GP:3}
'{Ep:Upall}
'{Caption:更新零件}
'{ControlTipText:遍历结构树并更新}
'{BackColor:}

Sub Upall()
   If Not CanExecute("ProductDocument") Then Exit Sub
    Dim part, doc
    For Each doc In CATIA.Documents
        If TypeName(doc) = "PartDocument" Then
            Set part = doc.part
            Exit For
        End If
    Next
'tosave =doc.saved
'if tosave =false then
    For Each doc In CATIA.Documents
      isupdated = True
      If TypeName(doc) = "PartDocument" Then
          isupdated = part.IsUpToDate(doc.part)
      ElseIf TypeName(doc) = "ProductDocument" Then
          isupdated = part.IsUpToDate(doc.Product)
      End If

    If Not isupdated Then
        On Error Resume Next
        doc.part.Update
        doc.Product.Update
        doc.Product.referenceprodcut.Parent.Update
    End If
    
    Next

End Sub

