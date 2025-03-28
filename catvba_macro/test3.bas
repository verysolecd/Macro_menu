Attribute VB_Name = "test3"
Sub CATMain()

Set rootPrd = CATIA.ActiveDocument.Product
Set oPrd = rootPrd.Products.item(4)


Set pdm = New class_PDM

arr = pdm.infoPrd(oPrd)








End Sub




Function count_me(oPrd)  '获取兄弟字典
     Dim i, oDict, QTy, pn
         QTy = 1
     On Error Resume Next
          If TypeOf oPrd.Parent Is Products Then    '若有父级产品'获取兄弟字典
                    Dim oParent: Set oParent = oPrd.Parent.Parent
                   Set oDict = CreateObject("Scripting.Dictionary")
                   For i = 1 To oParent.Products.Count
                          pn = oParent.Products.item(i).PartNumber
                          If oDict.Exists(pn) = True Then
                              oDict(pn) = oDict(pn) + 1
                          Else
                              oDict(pn) = 1
                          End If
                      Next
            QTy = oDict(oPrd.PartNumber)
          End If
    count_me = QTy
End Function
