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
On Error GoTo ErrorHandler
        Dim odata As Variant
           odata = xlm.extract_ary
           End If
'map 修改ary
      Dim iCols
    iCols = Array(0, 2, 4, 6, 8, 10, 12)
    Dim outputArr As Variant, temparr(1 To 6)
    
    ReDim outputArr(1 To UBound(odata, 1), 1 To UBound(iCols))
    
    For I = 1 To UBound(outputArr, 1)
        For j = 1 To UBound(outputArr, 2)
             outputArr(I, j) = ""
             If IsEmpty(odata(I, iCols(j))) = False Then
                outputArr(I, j) = odata(I, iCols(j))
             End If
             temparr(j) = outputArr(I, j)
        Next j
        
        Select Case I
            Case 1
            Case 2
            Call pdm.modatt(Prd2rv, temparr)
            Case Else
            Call pdm.modatt(children.item(I - currRow), temparr)
         End Select
        Next I
          Set Prd2rv = Nothing
       MsgBox "已经修改产品"
ErrorHandler:
    If Err.Number <> 0 Then
        Err.Clear
        MsgBox "程序错误：" & Err.Description, vbCritical
        Exit Sub
    End If
End Sub
