Attribute VB_Name = "OTH_SWScr"
'Attribute VB_Name = "OTH_EFF"
'{GP:6}
'{Ep:switchRefresh}
'{Caption: 屏幕更新}
'{ControlTipText:禁止屏幕更新以防止卡顿}
'{BackColor: }

Sub switchRefresh()

    CATIA.ActiveWindow.ActiveViewer.Update
    
End Sub
