Attribute VB_Name = "Module1"

Sub test22()

 '---显示管理
    
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
        
'--- 隐藏\显示

Visp.SetShow 0  '' 设置为可见
Visp.SetShow 1  '' 设置为不可见

    
'--颜色\线型

    Call Visp.SetRealColor(128, 64, 64, 1)
    Call Visp.SetRealOpacity(128, 1)
    Call Visp.SetRealWidth(1, 1)
    Call Visp.SetRealLineType(4, 1)

'    Set bdys = oPrt.bodies
'    Set bdy = getItem("Mini", bdys)
'    Set osel = CATIA.ActiveDocument.Selection
'    osel.Add bdy
'    osel.Delete
End Sub
