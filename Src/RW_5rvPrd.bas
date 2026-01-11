Attribute VB_Name = "RW_5rvPrd"
'{GP:1}
'{Ep:rvme}
'{Caption:修改产品属性}
'{ControlTipText:选择要被读取或修改的产品}
'{BackColor: }

Sub rvme()
     If pdm.CurrentProduct Is Nothing Then: MsgBox "请先选择产品，程序将退出": Exit Sub
        Dim currRow: currRow = 2
'---------遍历修改产品及子产品------
        Dim Prd2rv
        Set Prd2rv = pdm.CurrentProduct
            Prd2rv.ApplyWorkMode (3)
        Dim children: Set children = Prd2rv.Products
On Error GoTo errorhandler
        Dim odata As Variant: odata = xlm.extract_ary
'--------map 修改ary------
    Dim iCols: iCols = Array(0, 2, 4, 6, 8, 10, 12)
    Dim outputArr As Variant, temparr(1 To 6)
    ReDim outputArr(1 To UBound(odata, 1), 1 To UBound(iCols))
    
    
    
    For i = 1 To UBound(outputArr, 1) '-------遍历行
  '-------遍历获取X行要修改的数据
        For j = 1 To UBound(outputArr, 2) '-------遍历该行数组后输出一个一维数组作为要的修改参数
             outputArr(i, j) = ""  '-------初始化为空数组
             If IsEmpty(odata(i, iCols(j))) = False Then outputArr(i, j) = odata(i, iCols(j))
             temparr(j) = outputArr(i, j) '------
        Next j
    '-------遍历按行区分要修改的产品
        Select Case i
        Case 1
            Case 2  '第二行修改子总成
                Call pdm.modatt(Prd2rv, temparr)
            Case Else '其他修改子产品
                Call pdm.modatt(children.item(i - currRow), temparr)
         End Select
    Next i
          Set Prd2rv = Nothing
    MsgBox "已经修改产品"
errorhandler:
    If Err.Number <> 0 Then: Err.Clear: MsgBox "程序错误：" & Err.Description, vbCritical
        Exit Sub

End Sub
