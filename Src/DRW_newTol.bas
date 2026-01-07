Attribute VB_Name = "DRW_newTol"
Sub CATMain()


Set oDrw = CATIA.ActiveDocument
Set rtDrw = oDrw.DrawingRoot
Set Shts = rtDrw.Sheets
Set oSHT = Shts.item(1)
Set oVs = oSHT.Views
Set oView = oVs.ActiveView

Set ogdt = oView.GDTs.item(1) 'Add(1, 1, 20, 20, 10, "00")
tex = ogdt.GetReferenceNumber(1)

End Sub

