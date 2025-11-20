Attribute VB_Name = "OTH_designlog"
'Attribute VB_Name = "OTH_designlog"
'{GP:6}
'{Ep:designlog}
'{Caption:截图到文件夹}
'{ControlTipText:选择要被读取或修改的产品}
'{BackColor:16744703}

Option Explicit
Sub designlog()
    
    If Not KCL.CanExecute("ProductDocument") Then Exit Sub
    
    If pdm Is Nothing Then Set pdm = New class_PDM






End Sub






