Attribute VB_Name = "MDL_LayersMng"
'Attribute VB_Name = "MDL_LayersMng"
' 获得识别特征下的所有孔中心
'{GP:4}
'{EP:LayersMng}
'{Caption:当前层创建YZ向图纸}
'{ControlTipText: 设置只显示当前图层，然后创建YZ向图纸}
'{BackColor:12648447}
Private i
Sub LayersMng()
If Not CanExecute("partDocument,productdocument") Then Exit Sub
Set rootDoc = CATIA.ActiveDocument
Set rootprd = rootDoc.Product
    Call appFilterLayer(rootDoc)
    Call addDrw(rootprd)
 '---显示管理
'    '---图层管理
'    Dim layer: layer = CLng(0)
'    Dim layertype As CatVisLayerType
'    Dim Visp, osel
'
'    Set osel = CATIA.ActiveDocument.Selection
'    osel.Clear
'    osel.Add mbd
'    Set Visp = osel.VisProperties
'
'    Visp.GetLayer layertype, layer
'    If (layertype = catVisLayerNone) Then
'        layer = -1
'    End If
'    If (layertype = catVisLayerBasic) Then
'        MsgBox "layer =" & layer
'    End If
'        MsgBox "layer =" & layer
'        Visp.SetLayer catVisLayerBasic, 100
''--- 隐藏\显示
'
'Visp.SetShow 0  '' 设置为可见
'Visp.SetShow 1  '' 设置为不可见
'
'
''--颜色\线型 管理
'
'    Call Visp.SetRealColor(128, 64, 64, 1)
'    Call Visp.SetRealOpacity(128, 1)
'    Call Visp.SetRealWidth(1, 1)
'    Call Visp.SetRealLineType(4, 1)
'    Set bdys = oPrt.bodies
'    Set bdy = getItem("Mini", bdys)
'    Set osel = CATIA.ActiveDocument.Selection
'    osel.Add bdy
'    osel.Delete
'oDoc.CurrentFilter = "All visible"
End Sub
Sub appFilterLayer(oDoc)
Dim oSel
Set oSel = CATIA.ActiveDocument.Selection
 '---显示过滤器管理管理
 ily = ""
 ly = oDoc.CurrentLayer
If ly <> "None" Then
     Dim btn, bTitle, bResult
      imsg = "只显示当前图层还是您输入一个图层？" & vbCrLf & vbCrLf
      imsg = imsg & "选择 “是”: 只显示当前图层 " & vbCrLf
      imsg = imsg & "选择 “否”: 输入一个显示图层" & vbCrLf
      imsg = imsg & "选择 “取消”: 退出" & vbCrLf & vbCrLf
       btn = vbYesNo + vbExclamation
       bResult = MsgBox(imsg, btn, "bTitle")  ' Yes(6),No(7),cancel(2)
       Select Case bResult
           Case 2: Exit Sub '===选择“取消”====
           Case 7:  '===选择“否”====
                ipt = KCL.GetInput("请输入你想显示的图层，逗号分割")
                If VarType(ipt) = vbString Then
                    ipt = VBA.LCase(ipt)
                    ipt = Split(ipt, ",") '过滤器转数组
                End If
                fstr = ""
                For i = LBound(ipt) To UBound(ipt)
                    ily = "Layer= " & CLng(ipt(i)) & "+ " & ily
                Next i
                fstr = ""
                fstr = Left(ily, Len(ily) - 2)
                    If fstr <> "" Then
                        filterdef = fstr
                         filtername = "only_" & fstr & "_shown"
                           oDoc.CreateFilter filtername, filterdef
                           oDoc.CurrentFilter = filtername
                    End If
                Case 6  '===选择“是”====
                 oDoc.CurrentFilter = "Only current layer visible"
       End Select
End If
End Sub
Function addDrw(iprd)
    Dim docs As Documents
    Dim Shts As DrawingSheets
    Dim drwDoc As DrawingDocument
    Dim sht As DrawingSheet
    Dim oVs As DrawingViews
    Dim oV As DrawingView
    Dim ViewGen As DrawingViewGenerativeLinks
    Dim ViewGBH As DrawingViewGenerativeBehavior
    Set docs = CATIA.Documents
    Set drwDoc = docs.Add("Drawing")
    Set Shts = drwDoc.Sheets
    Set sht = Shts.item("Sheet.1")
    Set oVs = sht.Views
 xdis = 200
 i = 1
If iprd.Products.count < 1 Then
    On Error Resume Next
     Set oprt = iprd.ReferenceProduct.Parent.part
     If Not oprt Is Nothing Then
        Set oV = oVs.Add("AutomaticNaming")
            Set ViewGen = oV.GenerativeLinks
            Set ViewGBH = oV.GenerativeBehavior
                ViewGBH.Document = prd
                ViewGBH.DefineFrontView 0#, 1#, 0#, 0#, 0#, 1#
                oV.X = xdis * i
                oV.Y = 300
                oV.[Scale] = 1#
            ViewGBH.Update
            oV.Activate
            i = i + 1
        End If
    On Error Resume Next
Else
        For Each prd In iprd.Products
            Set oV = oVs.Add("AutomaticNaming")
            oV.Name = prd.PartNumber & "VIEW YZ"
            Set ViewGen = oV.GenerativeLinks
            Set ViewGBH = oV.GenerativeBehavior
                ViewGBH.Document = prd
                ViewGBH.DefineFrontView 0#, 1#, 0#, 0#, 0#, 1#
                oV.X = xdis * i
                oV.Y = 300
                oV.[Scale] = 1#
            ViewGBH.Update
            oV.Activate
            i = i + 1
        Next
End If
End Function
