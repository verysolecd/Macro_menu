Attribute VB_Name = "m45_freegPrd"
'Attribute VB_Name = "selPrd"
'{GP:4}
'{Ep:freegprd}
'{Caption:释放产品}
'{ControlTipText:将待操作产品清空}
'{BackColor:16744703}


Sub freegprd()
    Set gPrd = Nothing

    Set ProductObserver.CurrentProduct = gPrd ' 这会自动触发事件
    MsgBox "已清空待操作产品"
    Call clearall
End Sub



