Attribute VB_Name = "m43_readPrd"
'Attribute VB_Name = "selPrd"
'{gp:4}
'{Ep:readPrd}
'{Caption:读取属性}
'{ControlTipText:读取待操作产品}
'{BackColor:16744703}


Sub readPrd()
'excel处理和catia处理类初始化
Dim xlm, pdm
Set xlm = New Class_XLM
Set pdm = New class_PDM
'---------获取待修改产品
On Error Resume Next   
     If Not gprd Is Nothing Then
        gprd.ApplyWorkMode (3)
        Dim currRow: currRow = 2
'---------遍历修改产品及子产品
        Dim Prd2Read: Set Prd2Read = gprd
        xlm.inject_data currRow, pdm.infoPrd(Prd2Read)
        Dim children
        Set children = Prd2Read.Products
        For i = 1 To children.Count
         currRow = i + 2
         xlm.inject_data currRow, pdm.infoPrd(children.Item(i))
        Next
        Set Prd2Read = Nothing
    Else
     MsgBox "请先选择产品，程序将退出"
     Exit Sub
     End If
On Error GoTo 0
End Sub
