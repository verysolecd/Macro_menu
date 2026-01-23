Attribute VB_Name = "DRW_newTol"
Private Const mdlname As String = "DRW_newTol"
Sub newTol()


Set oDrw = CATIA.ActiveDocument
Set rtDrw = oDrw.DrawingRoot
Set shts = rtDrw.sheets
Set osht = shts.item(1)
Set oVs = osht.Views
Set oView = oVs.ActiveView

Set ogdt = oView.GDTs.item(1) 'Add(1, 1, 20, 20, 10, "00")
tex = ogdt.GetReferenceNumber(1)

End Sub

