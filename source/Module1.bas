Attribute VB_Name = "Module1"

Sub test22()

    Set odoc = CATIA.ActiveDocument
    Set oprt = CATIA.ActiveDocument.part
    Set lstPara = oprt.Parameters.RootParameterSet.ParameterSets.item("Part_info")
    Set lstbdys = lstPara.DirectParameters.item("iBodys")
    
  Set colls = lstbdys.valuelist
    Set odic = KCL.InitDic
    Set keeplst = KCL.InitDic
    For Each currobj In colls
        objkey = KCL.GetInternalName(currobj)
        If Not odic.Exists(objkey) Then
          odic(objkey) = 1
          keeplst(objkey) = 1
          End If
     Next
 
     For Each itm In colls
      colls.Remove itm.Name
        Next itm

    For Each bdy In oprt.bodies
           If keeplst.Exists(KCL.GetInternalName(bdy)) Then colls.Add bdy
    Next
           
    
'    '---图层管理
'    Dim layer: layer = CLng(0)
'    Dim layertype As CatVisLayerType
'    Dim Visp
'    Dim osel
'    Set osel = CATIA.ActiveDocument.Selection
'    osel.Clear
'    osel.Add mbd
'    Set Visp = osel.VisProperties
'    Visp.GetLayer layertype, layer
'    If (layertype = catVisLayerNone) Then
'        layer = -1
'    End If
'    If (layertype = catVisLayerBasic) Then
'        MsgBox "layer =" & layer
'    End If
'        MsgBox "layer =" & layer
'        Visp.SetLayer catVisLayerBasic, 100
        
 '---显示管理
 
    
 

'    Set bdys = oPrt.bodies
'    Set bdy = getItem("Mini", bdys)
'    Set osel = CATIA.ActiveDocument.Selection
'    osel.Add bdy
'    osel.Delete
End Sub
