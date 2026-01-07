VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Cls_WsEvt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private WithEvents pWb As Workbook
Attribute pWb.VB_VarHelpID = -1
Public Sub init(wbInstance)
    Set pWb = wbInstance
End Sub
Private Sub pWb_BeforeClose(Cancel As Boolean)
    ' 清除实例并释放对象
    Set gws = Nothing
    Set gwb = Nothing
     DoEvents
   Set xlAPP = Nothing
    Exit Sub
End Sub
