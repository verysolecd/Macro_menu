Attribute VB_Name = "m42_rvPrd"
'{GP:4}
'{Ep:rvme}
'{Caption:修改产品}
'{ControlTipText:选择要被读取或修改的产品}
'{BackColor: }

Sub rvme()

     If Not gPrd Is Nothing Then
        gPrd.ApplyWorkMode (3)
        Dim currRow: currRow = 2
        
'---------遍历修改产品及子产品   Set data =
        Dim Prd2rv
        Set Prd2rv = gPrd
        
        Dim odata As Variant
        odata = xlm.extract_data(2)

        
        Call pdm.modatt(Prd2rv, odata)
        Dim children
        Set children = Prd2rv.Products
        If children.Count > 0 Then
            For i = 1 To children.Count
            currRow = currRow + 1
            Dim sdata As Variant
            sdata = xlm.extract_data(currRow)
           Call pdm.modatt(children.item(i), sdata)
        Next
        End If
        Set Prd2rv = Nothing
        MsgBox "已经修改产品"
    Else
        MsgBox "请先选择产品，程序将退出"
        Exit Sub
     End If
    On Error GoTo 0


End Sub


Public Function extract_data(indRow)

    Dim iCols
    iCols = Array(0, 2, 4, 6, 8, 10, 12)

    
     Set ws = xlApp.ActiveSheet
     
    
    Dim temparr As Variant
    
    temparr = ws.Rows(indRow).Resize(1, 14).Value
    

    Dim outputArr As Variant
    Dim j As Long
    ReDim outputArr(1 To UBound(iCols))
    For j = 1 To UBound(iCols)
         outputArr(j) = ""
         If IsEmpty(temparr(1, iCols(j))) = False Then
         outputArr(j) = temparr(1, iCols(j))
         End If
    Next
    extract_data = outputArr ' 返回提取的数据
End Function




