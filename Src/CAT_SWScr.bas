Attribute VB_Name = "CAT_SWScr"

'{GP:7}
'{Ep:switchRefresh}
'{Caption: 屏幕更新}
'{ControlTipText:禁止屏幕更新以防止卡顿}
'{BackColor: }
Private Quick
Private Asmg, Vismg
Private Const mdlName As String = "CAT_SWScr"

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

Public Function setASM(ByVal Quick As Boolean)
    Dim btnCaption As String
    With CATIA
    If Quick Then
        '.DisableNewUndoRedoTransaction
        '.EnableNewUndoRedoTransaction
         .RefreshDisplay = False
            Asmg.AutoUpdateMode = 0 '0: catManualUpdate
            Vismg.Viz3DFixedAccuracy = 5
            btnCaption = "屏幕更新(关)"
    Else
        '.DisableNewUndoRedoTransaction
        '.EnableNewUndoRedoTransaction
        .RefreshDisplay = True
       Asmg.AutoUpdateMode = 1 '1: catAutomaticUpdate
        Vismg.Viz3DFixedAccuracy = 0.02
        btnCaption = "屏幕更新(开)"
    End If
    End With
    
    If Not A00_globalVar.g_Btn Is Nothing Then
        A00_globalVar.g_Btn.Caption = btnCaption
    End If
     Set A00_globalVar.g_Btn = Nothing
    setASM = Quick
    
    On Error Resume Next
    CATIA.ActiveWindow.ActiveViewer.Update
    On Error GoTo 0
End Function

