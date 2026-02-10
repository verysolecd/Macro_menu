Attribute VB_Name = "ASM_Updateme"
'{GP:3}
'{Ep:Upall}
'{Caption:更新零件}
'{ControlTipText:遍历结构树并更新}
Private Const mdlname As String = "ASM_Updateme"
Sub Upall()
   If Not CanExecute("ProductDocument,partdocument") Then Exit Sub
    Dim part, doc
    For Each doc In CATIA.Documents
        If TypeName(doc) = "PartDocument" Then Set part = doc.part: Exit For
    Next

On Error Resume Next
    For Each doc In CATIA.Documents
      isupdated = True
        If TypeName(doc) = "PartDocument" Then
            isupdated = part.IsUpToDate(doc.part)
        ElseIf TypeName(doc) = "ProductDocument" Then
            isupdated = part.IsUpToDate(doc.Product)
        End If
        If isupdated = False Then
            doc.part.Update
            doc.Product.Update
            doc.Product.referenceprodcut.Parent.Update
        End If
    Next
On Error GoTo 0
End Sub

