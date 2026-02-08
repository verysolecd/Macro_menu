Attribute VB_Name = "OTH_ivhideshow"
'Attribute VB_Name = "m64_hide&show"
'{GP:6}
'{Ep:setHideshow}
'{Caption:反选隐藏}
'{ControlTipText:反选并隐藏结构树}
'{BackColor:}

'标题格式为 %Title <Caption/Text>
'%Title 如何配置?
'------控件清单--------------------------------------------------
'控件格式为 %UI <ControlType> <ControlName> <Caption/Text>
' %UI Label lbL_jpzcs  键盘造车手出品
' %UI Button back2ori 恢复显示
' %UI Button allshow 显示所有产品
' %UI Button sel_child_show 显示选定产品
' %UI Button AsmHide_Plns 隐藏所有平面
' %UI Label lbL_4 ------
' %UI Button onlysel_hide  隐藏选定产品only
' %UI Button allhide 隐藏所有
' %UI Label lbL_5  '--以下针对零件--'
' %UI Button PrtHide_GS 隐藏根GSS
' %UI Button PrtHide_Skt 隐藏所有草图


Private msel, rdoc, rprd
Private showlst, hidelst, wantlst
Private orishowlst, orihidelst
Private Const mdlname As String = "OTH_ivhideshow"

Sub setHideshow()
   If Not KCL.CanExecute("ProductDocument,partdocument") Then Exit Sub
    If pdm Is Nothing Then Set pdm = New Cls_PDM
   initori
'==生成UItoolbar-===================
    Dim mapMdl: Set mapMdl = KCL.setBTNmdl(mdlname)
    Dim mapFunc As Object: Set mapFunc = KCL.setBTNFunc(mdlname)
    Set g_frm = Nothing:  Set g_frm = KCL.newFrm(mdlname)
    g_frm.ShowToolbar mdlname, mapMdl, mapFunc
End Sub
Private Sub initsel()
    Set rdoc = CATIA.ActiveDocument
    Set rprd = rdoc.Product
    Set msel = rdoc.Selection
End Sub
Private Sub Initlst()
    initsel
    Set wantlst = getwantlst
    Set showlst = getshowlst(rprd)
    Set hidelst = gethidelst(rprd)
End Sub
Private Sub initori()
    initsel
    Set orishowlst = getshowlst(rprd)
    Set orihidelst = gethidelst(rprd)
End Sub
Sub back2ori_click()
    On Error Resume Next
        hide_in_lst orihidelst
        show_in_lst orishowlst
    On Error GoTo 0
End Sub
Function getSubAsm()
    Set rdoc = CATIA.ActiveDocument
    Set rprd = rdoc.Product
    Set msel = rdoc.Selection
    Set oPrd = CATIA.ActiveDocument.Product
    Dim istr As String
    istr = "CATProductSearch.Product.Visibility=Hidden"
    SelQuery istr, oPrd
End Function

Sub PrtHide_GS_click()
    Set oprt = KCL.get_inwork_part
    If oprt Is Nothing Then Exit Sub
    Set HSF = oprt.HybridBodies
    Dim lst: Set lst = KCL.Initlst
    For Each itm In oprt.HybridBodies
        lst.Add itm
    Next
    hide_in_lst lst
End Sub

Sub PrtHide_Skt_click()
    initsel
    Set oprt = KCL.get_inwork_part
    If oprt Is Nothing Then Exit Sub
    msel.Clear
    Dim ss As String: ss = "Type=Sketch"
    SelQuery ss, oprt
    Dim lst: Set lst = KCL.Initlst
    For i = 1 To msel.count
       lst.Add msel.item(i).Value
'      lst.Add msel.item(i).LeafProduct
    Next
    
    hide_in_lst lst
    msel.Clear
End Sub
Sub AsmHide_Plns_click()
    Dim sel As Selection
    Set sel = CATIA.ActiveDocument.Selection
    sel.Clear
    sel.Search "Type=Plane,all"
    If sel.count > 0 Then
        sel.VisProperties.SetShow catVisPropertyNoShowAttr
    End If
    sel.Clear
End Sub

Sub PrtHide()
    Set oprt = KCL.get_inwork_part
' Dim istr As String: istr =
''Part Design'.Sketch&
''Generative Shape Design'.Sketch&
''Functional Molded Part'.Sketch
 
  filter = "(CATPrtSearch.BodyFeature.Visibility=Shown " & _
            "+ CATPrtSearch.OpenBodyFeature.Visibility=Shown" & _
            "+ CATPrtSearch.MMOrderedGeometricalSet.Visibility=Shown),sel"
            
            
         filter(1) = "(((CATStFreeStyleSearch.Plane + CATPrtSearch.Plane) + CATGmoSearch.Plane) + CATSpdSearch.Plane),all"
         filter(2) = "(((CATStFreeStyleSearch.AxisSystem + CATPrtSearch.AxisSystem) + CATGmoSearch.AxisSystem) + CATSpdSearch.AxisSystem),all"
         filter(3) = "((((((CATStFreeStyleSearch.Point + CAT2DLSearch.2DPoint) + CATSketchSearch.2DPoint) + CATDrwSearch.2DPoint) + CATPrtSearch.Point) + CATGmoSearch.Point) + CATSpdSearch.Point),all"
         filter(4) = "((((((CATStFreeStyleSearch.Curve + CAT2DLSearch.2DCurve) + CATSketchSearch.2DCurve) + CATDrwSearch.2DCurve) + CATPrtSearch.Curve) + CATGmoSearch.Curve) + CATSpdSearch.Curve),all"
         filter(5) = "(((CATStFreeStyleSearch.Surface + CATPrtSearch.Surface) + CATGmoSearch.Surface) + CATSpdSearch.Surface),all"
         filter(6) = "(((((((CATProductSearch.MfConstraint + CATStFreeStyleSearch.MfConstraint) + CATAsmSearch.MfConstraint) + CAT2DLSearch.MfConstraint) + CATSketchSearch.MfConstraint) + CATDrwSearch.MfConstraint) + CATPrtSearch.MfConstraint) + CATSpdSearch.MfConstraint),all"
 
 
    SelQuery istr, oPrd
    Set HSF = oprt.HybridBodies
    Dim lst: Set lst = KCL.Initlst
    For Each itm In oprt.HybridBodies
        lst.Add itm
    Next
    hide_in_lst lst
End Sub

Sub onlyselshow_click()
    initsel
    Initlst
    hide_in_lst showlst
    hide_in_lst hidelst
    show_in_lst wantlst
    showParent wantlst
End Sub

Sub sel_child_show_click()
    initsel
    Initlst
    hide_in_lst showlst
    hide_in_lst hidelst
    show_in_lst wantlst
    showChild wantlst
    showParent wantlst
End Sub

Sub onlysel_hide_click()
    initsel
    Initlst
    If msel Is Nothing Then Exit Sub
    On Error Resume Next
    show_in_lst showlst
    show_in_lst hidelst
    hide_in_lst wantlst
    On Error Resume Next
End Sub
Sub allshow_click()
     initsel
        Initlst
    On Error Resume Next
        show_in_lst orishowlst
        show_in_lst wantlst
        show_in_lst orihidelst
    On Error GoTo 0
End Sub
Sub allhide_click()
    On Error Resume Next
        hide_in_lst orishowlst
        hide_in_lst wantlst
        hide_in_lst orihidelst
    On Error GoTo 0
End Sub

Private Function SelQuery(iQuery As String, Optional ByVal iRange = Nothing)
  msel.Clear
  If iRange Is Nothing Then
    msel.Search iQuery & ",all"
  Else
     msel.Add iRange
     msel.Search iQuery & ",sel"
  End If
End Function

Function getshowlst(oPrd)
 Dim istr$: istr = "Assembly Design.Product.Visibility=Visible"
 Call SelQuery(istr, oPrd)
  Dim lst: Set lst = KCL.Initlst
   For i = 1 To msel.count
      lst.Add msel.item(i).LeafProduct
    Next
   Set getshowlst = lst
   msel.Clear
End Function

Function gethidelst(oPrd)
    Dim istr$: istr = "Assembly Design.Product.Visibility=Hidden"
    SelQuery istr, oPrd
    Dim lst: Set lst = KCL.Initlst
        For i = 1 To msel.count
        lst.Add msel.item(i).LeafProduct
    Next
        Set gethidelst = lst
    msel.Clear
End Function

Function getwantlst()
    Dim lst: Set lst = KCL.Initlst
    Set getwantlst = lst
If msel.Count2 > 0 Then
   For i = 1 To msel.count
      lst.Add msel.item(i).LeafProduct
    Next
      Set getwantlst = lst
      msel.Clear
End If
End Function
Sub hide_in_lst(lst)
    msel.Clear
    For Each itm In lst
            msel.Add itm
    Next
    msel.VisProperties.SetShow 1: msel.Clear
End Sub
Sub show_in_lst(lst)
    msel.Clear
    For Each itm In lst
            msel.Add itm
    Next
    msel.VisProperties.SetShow 0: msel.Clear
End Sub
Private Sub showParent(lst)
    Dim parentLst: Set parentLst = KCL.Initlst()
    Dim prd, parentPrd
    For Each prd In lst
        Set parentPrd = prd
        Do
            On Error Resume Next
            Set parentPrd = parentPrd.Parent
            If Err.Number <> 0 Then Err.Clear: Exit Do
            If TypeName(parentPrd) = "Product" Then
                 parentLst.Add parentPrd
            ElseIf TypeName(parentPrd) = "Products" Then
                 Set parentPrd = parentPrd.Parent
                 If TypeName(parentPrd) = "Product" Then parentLst.Add parentPrd
            Else
                 Exit Do ' 到达顶层或非 Product 结构
            End If
            On Error GoTo 0
        Loop
    Next
    If parentLst.count > 0 Then show_in_lst parentLst
End Sub

Sub showChild(lst)
    On Error Resume Next
    For Each itm In lst
       Set ilst = KCL.Initlst
     show_in_lst getchildlst(itm, ilst)
    Next
    On Error GoTo 0
End Sub

Function getchildlst(oPrd, lst)
    On Error Resume Next
      lst.Add oPrd
    For Each itm In oPrd.Products
        lst.Add itm
        getchildlst itm, lst
    Next
    Set getchildlst = lst
End Function

