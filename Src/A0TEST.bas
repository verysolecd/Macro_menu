Attribute VB_Name = "A0TEST"
   Sub test()


    Dim oSel As Selection
    Set oSel = CATIA.ActiveDocument.Selection
    

ss = oSel.item(1).Type
    
HybridShapeExtrude
    
   End Sub
