Attribute VB_Name = "m52_Cbom"
'Attribute VB_Name = "m5_Cbom"
'{GP:5}
'{Ep:cBom}
'{Caption:选择产品}
'{ControlTipText:选择要被读取或修改的产品}
'{BackColor:16744703}


Sub cBom()

    Dim xlm As New Class_XLM
    Dim pdm As New class_PDM
    Dim oPrd As Object
    ' 获取要处理的产品
    Set oPrd = pdm.catchPrd()
    If oPrd Is Nothing Then
        MsgBox "未选择产品"
        Exit Sub
    End If
    
    ' 获取BOM数据
    Dim bomArray As Variant
    bomArray = pdm.recurPrd(oPrd, 1)
    
    ' 一次性写入Excel
    Dim lastRow As Long
    lastRow = UBound(bomArray, 1)
  ws.Range(ws.Cells(2, 1), ws.Cells(lastRow + 1, UBound(bomArray, 2))).Value = bomArray
    
    ' 设置Excel可见
    xlApp.Visible = True
End Sub


      


