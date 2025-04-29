Attribute VB_Name = "m51_Cal_Mass"
'{GP:5}
'{Ep:Cal_Mass}
'{Caption:迭代重量}
'{ControlTipText:选择要被读取或修改的产品}
'{BackColor:16744703}

Sub Cal_Mass()
If gPrd Is Nothing Then
Call setgprd
End If

If pdm Is Nothing Then
    Set pdm = New class_PDM
End If

On Error Resume Next
Call pdm.Assmass(gPrd)
If Error.Number <> 0 Then
    MsgBox "重量已经计算"
Else
  MsgBox "程序错误：" & Err.Description, vbCritical
End If
On Error GoTo 0
End Sub


