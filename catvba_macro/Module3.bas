Attribute VB_Name = "Module3"
Sub pgupdate()
Set oDoc = CATIA.ActiveDocument
Set oShts = oDoc.Sheets
k = oShts.Count
On Error Resume Next
i = 1
For i = 1 To oShts.Count
    Dim oSHT As DrawingSheet
    Set oSHT = oShts.item(i)
    oSHT.Activate
        oo = straf1st(oSHT.Name, " ")
        oSHT.Name = "SH" & i & oo
    Set oviews = oSHT.Views
    Set oView = oviews.item("Background View")
    Set ots = oView.Texts
    Set Pg = ots.GetItem("dixxzhang")
    Pg.Text = "��" & i & "ҳ"
Set Pg = ots.GetItem("gongxxzhang")

  Pg.Text = "��" & k - 1 & "ҳ"
Next
On Error GoTo 0
End Sub


