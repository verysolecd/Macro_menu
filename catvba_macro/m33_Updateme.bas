Attribute VB_Name = "m33_Updateme"
'Attribute VB_Name = "m33_Updateme"
'{GP:3}
'{Ep:Upall}
'{Caption:更新零件}
'{ControlTipText:遍历结构树并更新}
'{BackColor:}

Sub Upall()

    Dim part
    Dim doc
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
          isupdated = part.isupdate(doc.part)
          ElseIf TypeName(doc) = "ProductDocument" Then
          isupdated = part.IsUpToDate(doc.product)
      End If


    If Not isupdated Then
        On Error Resume Next
        doc.part.Update
        doc.product.Update
        doc.product.referenceprodcut.Parent.Update
    End If
    
    Next

End Sub

