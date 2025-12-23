Attribute VB_Name = "DRW_updrw"
Sub DRA()
Set oDoc = CATIA.ActiveDocument
Set oShts = oDoc.Sheets
k = oShts.count
I = 5
    Dim oSHT As DrawingSheet
    Set oSHT = oShts.item(I)
    oSHT.Activate
        oo = straf1st(oSHT.Name, " ")
        oSHT.Name = "SH" & I & oo
    Set oViews = oSHT.Views
    Set oView = oViews.item("Background View")
   oView.Activate
    Set ots = oView.Texts
    Set Pg = ots.getItem("gongxxzhang")
    Pg.Text = "¹²" & k - 1 & "Ò³"
      oView.SaveEdition
    Set Pg = ots.getItem("dixxzhang")
    Pg.Text = "µÚ" & I & "Ò³"
    oView.SaveEdition
End Sub

  

