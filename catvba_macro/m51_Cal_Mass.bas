Attribute VB_Name = "m51_Cal_Mass"
'{GP:5}
'{Ep:Cal_Mass}
'{Caption:迭代重量}
'{ControlTipText:选择要被读取或修改的产品}
'{BackColor:16744703}

Sub Cal_Mass()
    If Not KCL.CanExecute("ProductDocument") Then Exit Sub

    If pdm Is Nothing Then
        Set pdm = New class_PDM
    End If
   On Error Resume Next
            If Not gPrd Is Nothing Then
                Call pdm.Assmass(gPrd)
            Else
                Call setgprd
                Err.Clear
                Call pdm.Assmass(gPrd)
            End If
            
   If Err.Number > 0 Then
        MsgBox "程序错误,请确认零件模板是否应用：" & Err.Description, vbCritical
   Else
            MsgBox "重量已计算"
    End If


End Sub
