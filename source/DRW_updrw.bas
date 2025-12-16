Attribute VB_Name = "DRW_updrw"
Sub DRA()
Set oDoc = CATIA.ActiveDocument
Set oShts = oDoc.Sheets
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
    Set Pg = ots.getItem("gongxxzhang")
    Pg.Text = "¹²" & k - 1 & "Ò³"
      oView.SaveEdition
    Set Pg = ots.getItem("dixxzhang")
    Pg.Text = "µÚ" & i & "Ò³"
    oView.SaveEdition
End Sub

  

