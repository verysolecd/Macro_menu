Attribute VB_Name = "RW_cMass"
'{GP:1}
'{Ep:Cal_Mass_m}
'{Caption:迭代重量}
'{ControlTipText:选择要被读取或修改的产品}
'{BackColor:}

Private Const mdlname As String = "BOM_cMass"
Sub Cal_Mass_m()
    If Not KCL.CanExecute("ProductDocument") Then Exit Sub
    If pdm Is Nothing Then Set pdm = New Cls_PDM
    On Error Resume Next
        If pdm.CurrentProduct Is Nothing Then Call setgprd: Err.Clear
        If Not pdm.CurrentProduct Is Nothing Then
            pdm.Assmass pdm.CurrentProduct
        End If
    If Err.Number > 0 Then
        MsgBox "程序错误,请确认零件模板是否应用：" & Err.Description, vbCritical
   Else
        MsgBox "重量已计算"
   End If
   On Error GoTo 0
End Sub
Sub Cal_Mass2()
    If pdm Is Nothing Then Set pdm = New Cls_PDM
   On Error Resume Next
            If pdm.CurrentProduct Is Nothing Then Call setgprd
            Err.Clear
            If Not pdm.CurrentProduct Is Nothing Then pdm.Assmass (pdm.CurrentProduct)
    On Error GoTo 0
End Sub

