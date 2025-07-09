Attribute VB_Name = "m33_Updateme"
'Attribute VB_Name = "m30_NewPn"
'{GP:3}
'{Ep:Upall}
'{Caption:更新所有零件}
'{ControlTipText:为所有零件号增加项目前缀}
'{BackColor:}
Sub main()

    Dim part
    Dim doc
    For Each doc In CATIA.Documents
        If TypeName(odc) = "PartDocument" Then
            Set part = doc.part
            Exit For
        End If
    Next


'tosave =doc.saved
'if tosave =false then
    dim isupdated

    For Each doc In CATIA.Documents
      isupdated = True
      If TypeName(doc) = "PartDocument" Then
          isupdated = part.isupdate(doc.part)
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

