Attribute VB_Name = "m40_setgprd"
'{GP:4}
'{Ep:setgprd}
'{Caption:选择产品}
'{ControlTipText:选择要被读取或修改的产品}
'{BackColor:16744703}

Sub setgprd()


    If Not CanExecute("ProductDocument") Then Exit Sub


    If pdm Is Nothing Then
        Set pdm = New class_PDM
    End If

    
    Dim oprd
    Set gPrd = pdm.defgprd()
    Set ProductObserver.CurrentProduct = gPrd ' 这会自动触发事件
         
        If Not gPrd Is Nothing Then
           imsg = "你选择的产品是" & gPrd.PartNumber
            MsgBox imsg
        Else
             MsgBox "已退出，程序将结束"
        End If
End Sub
