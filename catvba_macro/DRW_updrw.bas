Attribute VB_Name = "DRW_updrw"
Sub DRA()
Set oDoc = CATIA.ActiveDocument
Set oShts = oDoc.Sheets
k = oShts.Count
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
    Pg.Text = "��" & k - 1 & "ҳ"
      oView.SaveEdition
    Set Pg = ots.GetItem("dixxzhang")
    Pg.Text = "��" & i & "ҳ"
    oView.SaveEdition
End Sub

  

