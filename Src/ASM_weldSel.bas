Attribute VB_Name = "ASM_weldSel"
'Attribute VB_Name = "weldSel"
'{GP:3}
'{Ep:createWd}
'{Caption:焊接命名}
'{ControlTipText:选择被连接的元素后按零件号生成点焊几何图形集}
'{BackColor:}
Option Explicit
Private Const mdlname As String = "ASM_weldSel"
Sub createWd()
 If Not KCL.CanExecute("Productdocument,PartDocument") Then Exit Sub
    On Error Resume Next
        Dim oDoc: Set oDoc = CATIA.ActiveDocument
        Dim workPrtDoc: Set workPrtDoc = KCL.get_workPartDoc
        Dim oprt: Set oprt = Nothing: Set oprt = workPrtDoc.part
    Err.Clear
    On Error GoTo 0
If IsNothing(oprt) Then: MsgBox "No activated Part for welding info": Exit Sub
Dim osel: Set osel = oDoc.Selection: osel.Clear
Dim i, j, oPn, itm, itp, odict, wdGeoname: wdGeoname = ""
Set odict = KCL.InitDic
Set osel = CATIA.ActiveDocument.Selection: osel.Clear
    If osel.count = 0 Then Set osel = KCL.Selectmulti("请选择被点焊对象")  ', "Body,Product"
    For i = 1 To osel.count
        On Error Resume Next
             Set itp = osel.item(i).LeafProduct
             If Not itp Is Nothing Then
                oPn = itp.partNumber 'Parent.Product.
                If Not odict.Exists(oPn) Then odict.Add oPn, 1
             End If
         On Error GoTo 0
    Next i
 
' KCL.showdict odict
 Dim arrKeys: arrKeys = odict.keys
 If UBound(arrKeys) < 1 Or UBound(arrKeys) > 2 Then
    MsgBox "点焊层数错误"
    Exit Sub
 End If
' 排序
Dim tmp
For i = 0 To UBound(arrKeys) - 1
    For j = i + 1 To UBound(arrKeys)
        If StrComp(arrKeys(i), arrKeys(j), vbTextCompare) > 0 Then
            tmp = arrKeys(i)
            arrKeys(i) = arrKeys(j)
            arrKeys(j) = tmp
        End If
    Next j
Next i

' 排序后的 keys 拼接
Dim istr
For i = 0 To UBound(arrKeys)
    If i = 0 Then
        istr = arrKeys(i)
    Else
        istr = istr & "--" & arrKeys(i)
    End If
Next i
wdGeoname = "SotWeld_" & istr
Dim colls, og
    Set colls = oprt.HybridBodies
    Set og = colls.Add(): og.Name = wdGeoname
End Sub
 

