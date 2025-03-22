
Private WithEvents pWb As Workbook

Public Sub Init(wbInstance)
    Set pWb = wbInstance 
End Sub

Private Sub pWb_BeforeClose(Cancel As Boolean)
    Set gwb = Nothing
    Set gws = Nothing
    Set xlApp = Nothing
End Sub
