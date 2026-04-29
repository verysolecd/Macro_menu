Attribute VB_Name = "Module1"

Sub test22()


    Dim msg As String
    msg = "请选择产品"
    Dim prod As Product

  Set prod = KCL.SelectItem(msg, "Product")
  
  
    Set oPrt = CATIA.ActiveDocument.part

    Set bdys = oPrt.bodies
    Set bdy = getItem("Mini", bdys)
    Set osel = CATIA.ActiveDocument.Selection
    osel.Add bdy
    osel.Delete

End Sub

Function getItem(iName, colls)
 Dim itm ' 正确声明数组
    Set itm = Nothing
    On Error Resume Next
        Set itm = colls.item(iName)
            Err.Clear
            Err.Number = 0
    On Error GoTo 0
   Set getItem = itm
    Set itm = Nothing
End Function

