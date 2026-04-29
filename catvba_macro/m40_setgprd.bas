Attribute VB_Name = "m40_setgprd"
'{GP:4}
'{Ep:setgprd}
'{Caption:选择产品}
'{ControlTipText:选择要被读取或修改的产品}
'{BackColor:16744703}

Sub setgprd()

    If pdm Is Nothing Then
        Set pdm = New class_PDM
    End If
    
    Dim oprd
    
    imsg = MsgBox("“是”选择产品，“否”读取根产品，“取消退出”", vbYesNoCancel + vbDefaultButton2, "请选择产品")

        Select Case imsg
            Case 7 '===选择“否”====
                Set oprd = CATIA.ActiveDocument.Product
            Case 2:
                Exit Sub '===选择“取消”====
            Case 6  '===选择“是”,进行产品选择====
                On Error Resume Next
                    Set oprd = pdm.selPrd()
                    If Err.Number <> 0 Then
                        Err.Clear
                        Exit Sub
                    End If
                On Error GoTo 0
        End Select
        
         Set gPrd = oprd
         Set oprd = Nothing
         
        If Not gPrd Is Nothing Then
           imsg = "你选择的产品是" & gPrd.PartNumber
            MsgBox imsg
        Else
             MsgBox "已退出，程序将结束"
        End If
End Sub
