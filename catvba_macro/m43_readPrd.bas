Attribute VB_Name = "m43_readPrd"
'Attribute VB_Name = "selPrd"
'{gp:4}
'{Ep:readPrd}
'{Caption:读取属性}
'{ControlTipText:选择要被读取或修改的产品}
'{BackColor:16744703}

Sub readPrd()

'excel处理和catia处理类初始化
    Dim xlm, pdm
    Set xlm = New Class_XLM
    Set pdm = New class_PDM
    
'---------获取待修改产品
    On Error Resume Next
        Set gprd = pdm.catchPrd()
        If gprd Is Nothing Then
            MsgBox "未选择产品"
        End If
    On Error GoTo 0
    
    Dim currRow: currRow = 2
 '---------遍历修改产品及子产品
    Dim Prd2Read: Set Prd2Read = gprd
        Prd2Read.ApplyWorkMode (3)
        xlm.inject_data currRow, pdm.infoPrd(Prd2Read), "rv"
        
    Dim children
    Set children = Prd2Read.Products
        For i = 1 To children.Count
         currRow = i + 2
         xlm.inject_data currRow, pdm.infoPrd(children.Item(i)), "rv"
        Next
    Set Prd2Read = Nothing
    xlm.alapp.Visible = True
        
    
'ErrHandler:
'    Select Case Err.Number
'    Case 429 ' CATIA未运行
'    MsgBox "无法连接CATIA，请确保程序已运行", vbCritical
'    Case Else
'    MsgBox "错误 " & Err.Number & ": " & Err.Description, vbCritical
'    End Select
    
End Sub

