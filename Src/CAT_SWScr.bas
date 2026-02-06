Attribute VB_Name = "CAT_SWScr"

'{GP:7}
'{Ep:switchRefresh}
'{Caption: 屏幕更新}
'{ControlTipText:禁止屏幕更新以防止卡顿}
'{BackColor: }
Private Quick
Private Asmg, Vismg
Private Const mdlname As String = "CAT_SWScr"

Sub switchRefresh()

      Dim setcls:  Set setcls = CATIA.SettingControllers
    Set Asmg = setcls.item("CATAsmGeneralSettingCtrl")
   Set Vismg = setcls.item("CATVizVisualizationSettingCtrl")
On Error Resume Next
    CATIA.ActiveWindow.ActiveViewer.Update
On Error GoTo 0
     Quick = IIf(Vismg.Viz3DFixedAccuracy = 5, True, False)
    setASM Not Quick
On Error Resume Next
    CATIA.ActiveWindow.ActiveViewer.Update
On Error GoTo 0
End Sub

