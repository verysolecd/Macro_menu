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
' %UI Label lbL_4 ------
' %UI Button revertselShow  隐藏选定产品only
' %UI Button allhide 隐藏所有
' %UI Label lbL_5  '--针对零件--'
' %UI Button PrtHide_GS 隐藏根GSS
' %UI Button PrtHide_Skt 隐藏所有草图
Private msel, rdoc, rprd
Private showlst, hidelst, wantlst
Private orishowlst, orihidelst
Private Const mdlname As String = "OTH_ivhideshow"

Sub setHideshow()
If Not CanExecute("ProductDocument") Then Exit Sub
    If pdm Is Nothing Then Set pdm = New Cls_PDM
   initori
'==生成UItoolbar-===================
    Dim mapmdl: Set mapmdl = KCL.setBTNmdl(mdlname)
    Dim mapFunc As Object: Set mapFunc = KCL.setBTNFunc(mdlname)
    Set g_Frm = Nothing:  Set g_Frm = KCL.newFrm(mdlname)
    g_Frm.ShowToolbar mdlname, mapmdl, mapFunc
End Sub
Sub initsel()


End Sub
Sub PrtHide_GS_click()
    Set oPrt = KCL.get_inwork_part
    Set HSF = oPrt.HybridBodies
    Dim lst: Set lst = KCL.Initlst
    For Each itm In oPrt.HybridBodies
        lst.Add itm
    Next
    hide_in_lst lst
End Sub

Sub PrtHide_Skt_click()
    Set oPrt = KCL.get_inwork_part
    
 Dim istr As String: istr =
 
'Part Design'.Sketch&
'Generative Shape Design'.Sketch&
'Functional Molded Part'.Sketch
 
  filter = "(CATPrtSearch.BodyFeature.Visibility=Shown " & _
            "+ CATPrtSearch.OpenBodyFeature.Visibility=Shown" & _
            "+ CATPrtSearch.MMOrderedGeometricalSet.Visibility=Shown),sel"
            
            
         filter(1) = "(((CATStFreeStyleSearch.Plane + CATPrtSearch.Plane) + CATGmoSearch.Plane) + CATSpdSearch.Plane),all"
         filter(2) = "(((CATStFreeStyleSearch.AxisSystem + CATPrtSearch.AxisSystem) + CATGmoSearch.AxisSystem) + CATSpdSearch.AxisSystem),all"
         filter(3) = "((((((CATStFreeStyleSearch.Point + CAT2DLSearch.2DPoint) + CATSketchSearch.2DPoint) + CATDrwSearch.2DPoint) + CATPrtSearch.Point) + CATGmoSearch.Point) + CATSpdSearch.Point),all"
         filter(4) = "((((((CATStFreeStyleSearch.Curve + CAT2DLSearch.2DCurve) + CATSketchSearch.2DCurve) + CATDrwSearch.2DCurve) + CATPrtSearch.Curve) + CATGmoSearch.Curve) + CATSpdSearch.Curve),all"
         filter(5) = "(((CATStFreeStyleSearch.Surface + CATPrtSearch.Surface) + CATGmoSearch.Surface) + CATSpdSearch.Surface),all"
         filter(6) = "(((((((CATProductSearch.MfConstraint + CATStFreeStyleSearch.MfConstraint) + CATAsmSearch.MfConstraint) + CAT2DLSearch.MfConstraint) + CATSketchSearch.MfConstraint) + CATDrwSearch.MfConstraint) + CATPrtSearch.MfConstraint) + CATSpdSearch.MfConstraint),all"
 
 
 Call SelectAll(istr, oprd)
    
    Set HSF = oPrt.HybridBodies
    Dim lst: Set lst = KCL.Initlst
    For Each itm In oPrt.HybridBodies
        lst.Add itm
    Next
    
    
    hide_in_lst lst
End Sub


Function getSubAsm()
 Set rdoc = CATIA.ActiveDocument
    Set rprd = rdoc.Product
    Set msel = rdoc.Selection
Set oprd = CATIA.ActiveDocument.Product
Dim istr As String
istr = "CATProductSearch.Product.Visibility=Hidden"

SelectAll istr, oprd




End Function



Sub onlyselshow_click()
    Initlst
    hide_in_lst showlst
    hide_in_lst hidelst
    show_in_lst wantlst
    showParent wantlst
End Sub
Sub sel_child_show_click()
    Initlst
    hide_in_lst showlst
    hide_in_lst hidelst
    show_in_lst wantlst
    showChild wantlst
    showParent wantlst
End Sub

Sub revertselshow_click()
    Initlst
    '====开始隐藏
    show_in_lst showlst
    show_in_lst hidelst
    hide_in_lst wantlst
    ' show_in_lst showlst
    'show_in_lst hidelst
    '====开始恢复显示
End Sub
Sub allshow_click()
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

Sub back2ori_click()
On Error Resume Next
    hide_in_lst orihidelst
    show_in_lst orishowlst
On Error GoTo 0
End Sub

Private Sub Initlst()
    Set rdoc = CATIA.ActiveDocument
    Set rprd = rdoc.Product
    Set msel = rdoc.Selection
    Set wantlst = getwantlst
    Set showlst = getshowlst(rprd)
    Set hidelst = gethidelst(rprd)
End Sub

Sub initori()
    Set rdoc = CATIA.ActiveDocument
    Set rprd = rdoc.Product
    Set msel = rdoc.Selection
    Set orishowlst = getshowlst(rprd)
    Set orihidelst = gethidelst(rprd)
End Sub
Private Function SelectAll(iQuery As String, ByVal iRange)
  msel.Clear
  msel.Add iRange
  msel.Search iQuery & ",sel"
End Function
Function getshowlst(oprd)
 Dim istr As String: istr = "Assembly Design.Product.Visibility=Visible"
 Call SelectAll(istr, oprd)
  Dim lst: Set lst = KCL.Initlst
   For i = 1 To msel.count
      lst.Add msel.item(i).LeafProduct
    Next
   Set getshowlst = lst
   msel.Clear
End Function
Function gethidelst(oprd)
 Dim istr As String: istr = "Assembly Design.Product.Visibility=Hidden"
 Call SelectAll(istr, oprd)
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

Function getchildlst(oprd, lst)
    On Error Resume Next
      lst.Add oprd
    For Each itm In oprd.Products
        lst.Add itm
        getchildlst itm, lst
    Next
    Set getchildlst = lst
End Function

