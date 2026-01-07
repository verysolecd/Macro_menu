Attribute VB_Name = "RW_5rvPrd"
'{GP:1}
'{Ep:rvme}
'{Caption:修改产品属性}
'{ControlTipText:选择要被读取或修改的产品}
'{BackColor: }

Sub rvme()
     If gPrd Is Nothing Then
        MsgBox "请先选择产品，程序将退出"
        Exit Sub
     Else
        Dim currRow: currRow = 2
'---------遍历修改产品及子产品   Set data =
        Dim Prd2rv: Set Prd2rv = gPrd
        Dim children
                Set children = Prd2rv.Products
              Prd2rv.ApplyWorkMode (3)
On Error GoTo errorhandler
        Dim odata As Variant
           odata = xlm.extract_ary
           End If
'map 修改ary
      Dim iCols
    iCols = Array(0, 2, 4, 6, 8, 10, 12)
    Dim outputArr As Variant, temparr(1 To 6)
    
    ReDim outputArr(1 To UBound(odata, 1), 1 To UBound(iCols))
    
    For i = 1 To UBound(outputArr, 1)
        For j = 1 To UBound(outputArr, 2)
             outputArr(i, j) = ""
             If IsEmpty(odata(i, iCols(j))) = False Then
                outputArr(i, j) = odata(i, iCols(j))
             End If
             temparr(j) = outputArr(i, j)
        Next j
        
        Select Case i
            Case 1
            Case 2
            Call pdm.modatt(Prd2rv, temparr)
            Case Else
            Call pdm.modatt(children.item(i - currRow), temparr)
         End Select
        Next i
          Set Prd2rv = Nothing
       MsgBox "已经修改产品"
errorhandler:
    If Err.Number <> 0 Then
        Err.Clear
        MsgBox "程序错误：" & Err.Description, vbCritical
        Exit Sub
    End If
End Sub
