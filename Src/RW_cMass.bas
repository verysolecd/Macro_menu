Attribute VB_Name = "RW_cMass"
'{GP:1}
'{Ep:Cal_Mass_m}
'{Caption:迭代重量}
'{ControlTipText:选择要被读取或修改的产品}
'{BackColor:}


'  %UI Label  lblInfo  请选择操作：
'  %UI Button btna  更新重量
'  %UI Button btnb 更新LV2重量

Private Const mdlname As String = "RW_cMass"


Sub Cal_Mass_m()
    If Not KCL.CanExecute("ProductDocument") Then Exit Sub
    If pdm Is Nothing Then Set pdm = New Cls_PDM
    '==生成UItoolbar-===================
    Dim mapmdl: Set mapmdl = KCL.setBTNmdl(mdlname)
    Dim mapFunc As Object: Set mapFunc = KCL.setBTNFunc(mdlname)  'btnname_click
    Set g_Frm = Nothing:  Set g_Frm = KCL.newFrm(mdlname)
    g_Frm.ShowToolbar mdlname, mapmdl, mapFunc

End Sub
Sub btna_click()
 On Error Resume Next
       Cal_Mass
   If Err.Number > 0 Then
        MsgBox "程序错误,请确认零件模板是否应用：" & Err.Description, vbCritical
   Else
        MsgBox "重量已计算"
   End If
     On Error GoTo 0
End Sub

Sub btnb_click()
    L2Mass
    MsgBox "已经执行"
End Sub

Sub Cal_Mass()
   If pdm.CurrentProduct Is Nothing Then Call setgprd: Err.Clear
        If Not pdm.CurrentProduct Is Nothing Then
        Set oprd = pdm.CurrentProduct
            pdm.Assmass oprd
        End If
End Sub

Sub L2Mass()
   If pdm.CurrentProduct Is Nothing Then Call setgprd: Err.Clear
   Set oprd = pdm.CurrentProduct
   Call LV2_Mass(oprd, 1)
    Set oprd = Nothing
End Sub
Function LV2_Mass(oprd, Lv)
        If Lv <= 3 Then
                Set children = oprd.Products
                If children.count > 0 Then
                    For i = 1 To children.count
                        Call LV2_Mass(children.item(i), Lv + 1)
                        total = total + children.item(i).ReferenceProduct.UserRefProperties.item("Mass").value
                    Next
                        oprd.ReferenceProduct.UserRefProperties.item("Mass").value = total
                Else
                        total = oprd.ReferenceProduct.UserRefProperties.item("Mass").value
                End If
        End If
End Function
