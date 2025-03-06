Attribute VB_Name = "m43_readPrd"
'Attribute VB_Name = "selPrd"
'{gp:4}
'{Ep:readPrd}
'{Caption:读取属性}
'{ControlTipText:选择要被读取或修改的产品}
'{BackColor:16744703}

Sub readPrd()
    Dim xlm, pdm, ws
    Set xlm = New Class_XLM
    Set pdm = New class_PDM
    Set ws = gws
    
'---------获取待修改产品
    On Error resume Next
        Set g_Prd2wt = pdm.catchPrd()
        If g_Prd2wt Is Nothing Then
            MsgBox "未选择产品"
        End If
        if Err.Number <> 0 Then
            msg box "未选择产品"
    On Error GoTo 0
    On Error resume Next
    Dim currRow: currRow = 2
 '---------遍历修改产品及子产品
    Dim Prd2Read: Set Prd2Read = g_Prd2wt
        xlm.inject_data currRow, pdm.infoPrd(Prd2Read), "rv"
        
    Dim children
    Set children = Prd2Read.Products
        For i = 1 To children.Count
         currRow = i + 2
         xlm.inject_data currRow, pdm.infoPrd(children.Item(i)), "rv"
        Next
    Set Prd2Read = Nothing
     On Error GoTo 0   
    
ErrHandler:
    Select Case Err.Number
    Case 429 ' CATIA未运行
    MsgBox "无法连接CATIA，请确保程序已运行", vbCritical
    Case Else
    MsgBox "错误 " & Err.Number & ": " & Err.Description, vbCritical
    End Select
    
End Sub
