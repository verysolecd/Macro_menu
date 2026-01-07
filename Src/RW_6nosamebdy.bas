Attribute VB_Name = "RW_6nosamebdy"
'Attribute VB_Name = "RW_6nosamebdy"
'{GP:1}
'{Ep:nosamebdy}
'{Caption:计重实体去重}
'{ControlTipText:将零件计算重量的实体清单中重复的实体删除}
'{BackColor:1229803}
Option Explicit

Sub nosamebdy()
 If Not KCL.CanExecute("ProductDocument,PartDocument") Then Exit Sub

  Dim iprd, oprt
 If KCL.checkDocType("PartDocument") Then
    Set oprt = CATIA.ActiveDocument.part
    Call nosamebdy_bdylst(oprt)
    Set pdm = Nothing
    Exit Sub
 End If
  If pdm Is Nothing Then Set pdm = New Cls_PDM
   
  Set allPN = KCL.InitDic(vbTextCompare): allPN.RemoveAll  'allPn 是全局变量，不需要传递
 Set iprd = pdm.getiPrd()
 If Not iprd Is Nothing Then
     On Error Resume Next
            Call nosamebdy_prds(iprd)
            allPN.RemoveAll
       If Error.Number = 0 Then
                MsgBox "已删除重复实体"
          Else
                MsgBox "错误，可能有实体未删除"
                 End If
      On Error GoTo 0
   Else
    MsgBox "没有产品或零件，将退出"
 End If
 
 allPN.RemoveAll
End Sub
Sub nosamebdy_prds(oprd)
    Dim Product
        If allPN.Exists(oprd.PartNumber) = False Then
            allPN(oprd.PartNumber) = 1
            Call nosamebdy_prd(oprd)
        End If
    If oprd.Products.count > 0 Then
            For Each Product In oprd.Products
                Call nosamebdy_prds(Product)
             Next
    End If
End Sub
Public Sub nosamebdy_prd(oprd)
    Dim colls, oprt

    On Error Resume Next
         Set oprt = oprd.ReferenceProduct.Parent.part
        If Err.Number <> 0 Then
            Err.Clear
            Set oprt = Nothing
        End If
    On Error GoTo 0
    
    If Not oprt Is Nothing Then
        On Error Resume Next
                Call nosamebdy_bdylst(oprd)
                Err.Clear
        On Error GoTo 0
    End If
End Sub
Sub nosamebdy_bdylst(oprt)
    Dim lstPara, lstbdys, colls, oDic, keeplst, currobj, objkey, itm, bdy
    Set lstPara = oprt.Parameters.RootParameterSet.ParameterSets.item("Part_info")
    Set lstbdys = lstPara.DirectParameters.item("iBodys")
  Set colls = lstbdys.valuelist
    Set oDic = KCL.InitDic
    Set keeplst = KCL.InitDic
    For Each currobj In colls
        objkey = KCL.GetInternalName(currobj)
        If Not oDic.Exists(objkey) Then
          oDic(objkey) = 1
          keeplst(objkey) = 1
          End If
     Next
     For Each itm In colls
      colls.Remove itm.Name
        Next itm
    For Each bdy In oprt.bodies
           If keeplst.Exists(KCL.GetInternalName(bdy)) Then colls.Add bdy
    Next
End Sub


