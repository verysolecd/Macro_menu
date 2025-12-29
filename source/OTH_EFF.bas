Attribute VB_Name = "OTH_EFF"
'Attribute VB_Name = "OTH_EFF"
'{GP:6}
'{Ep:CATMain}
'{Caption: ÆÁÄ»Ë¢ÐÂ}
'{ControlTipText: ¿ØÖÆÆÁÄ»Ë¢ÐÂÒÔ·ÀÖ¹ÉÁË¸}
'{BackColor: }

Sub CATMain()
          bol = CATIA.RefreshDisplay
    If bol = False Then
                CATIA.RefreshDisplay = True
     Else
                CATIA.RefreshDisplay = False
    End If
          bol = CATIA.RefreshDisplay
'    Set setcls = CATIA.SettingControllers
'    Set Asmg = setcls.item("CATAsmGeneralSettingCtrl")
'
'    Asmg.AutoUpdateMode = 0  '0: catManualUpdate /
'    ''Asmg.AutoUpdateMode = 1  '1: catAutomaticUpdate
'
'    Set Vsmg = setcls.item("CATVizVisualizationSettingCtrl")
'    Vsmg.Viz3DFixedAccuracy = 1
'    Vsmg.Viz3DFixedAccuracy = 0.1
    
    
    With CATIA
    
    '.DisableNewUndoRedoTransaction
    '.EnableNewUndoRedoTransaction
    '.RefreshDisplay = False
    
    End With
    
    
    With CATIA
    
    
    '.DisableNewUndoRedoTransaction
    '.EnableNewUndoRedoTransaction
    
      '.BeginURConcatenation
        '.StopURConcatenation
    
    '.RefreshDisplay = False
    
    End With




End Sub