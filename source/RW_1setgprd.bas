Attribute VB_Name = "RW_1setgprd"
'{GP:1}
'{Ep:setgprd}
'{Caption:选择产品}
'{ControlTipText:选择要被读取或修改的产品}
'{BackColor:16744703}

Sub setgprd()
    If Not CanExecute("ProductDocument") Then Exit Sub
    If pdm Is Nothing Then
        Set pdm = New Cls_PDM
    End If

    Set gPrd = pdm.getiPrd()
    Set pdm.CurrentProduct = gPrd ' 这会自动触发事件

        If Not gPrd Is Nothing Then
           imsg = "你选择的产品是" & gPrd.PartNumber
            MsgBox imsg
        Else
             MsgBox "已退出，程序将结束"
        End If
End Sub
