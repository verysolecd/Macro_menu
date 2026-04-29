VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Ws_Evt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private WithEvents pWb As Workbook
Attribute pWb.VB_VarHelpID = -1
Public Sub Init(wbInstance)
    Set pWb = wbInstance
End Sub
Private Sub pWb_BeforeClose(Cancel As Boolean)
    ' 清除实例并释放对象
    Set gwb = Nothing
    Set gws = Nothing
    Set xlAPP = Nothing
End Sub
