Attribute VB_Name = "RW_2freegPrd"
'{GP:1}
'{Ep:freegprd}
'{Caption:释放产品}
'{ControlTipText:将待操作产品清空}
'{BackColor:16744703}

Private Const mdlname As String = "RW_2freegPrd"
Sub freegprd()
    Set pdm.CurrentProduct = Nothing
    MsgBox "已清空待操作产品"
    Call clearall
End Sub
