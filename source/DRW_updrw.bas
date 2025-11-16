Attribute VB_Name = "DRW_updrw"
Sub DRA()
Set odoc = catia.ActiveDocument
Set oShts = odoc.Sheets
k = oShts.count
i = 5
    Dim oSHT As DrawingSheet
    Set oSHT = oShts.item(i)
    oSHT.Activate
        oo = straf1st(oSHT.Name, " ")
        oSHT.Name = "SH" & i & oo
    Set oviews = oSHT.Views
    Set oView = oviews.item("Background View")
   oView.Activate
    Set ots = oView.Texts
    Set Pg = ots.GetItem("gongxxzhang")
    Pg.Text = "¹²" & k - 1 & "Ò³"
      oView.SaveEdition
    Set Pg = ots.GetItem("dixxzhang")
    Pg.Text = "µÚ" & i & "Ò³"
    oView.SaveEdition
End Sub

  

