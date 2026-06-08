Attribute VB_Name = "MDL_eleRename"
'{GP:4}
'{EP:eleRename}
'{Caption:元素重命名}
'{ControlTipText: 提示选择后将实体或几何图形集下的元素按顺序重命名}
'{BackColor: }
'' Type of feature
'= 0 , Unknown
'= 1 , Point
'= 2 , Curve
'= 3 , Line
'= 4 , Circle
'= 5 , Surface
'= 6 , Plane
'= 7 , Solid, Volume
'----------弹窗信息=----------------------------------
' %UI Label lbL_jpzcs  键盘造车手出品
' %UI txtbox txtname  Vbnull
' %UI Button bodyReName  实体重命名
' %UI Button hybodyReName shape重命名
' %UI Button HBReName  几何图形集重命名
' %UI Button btncancel  取消

Option Explicit
Private mWD
' ==================== 模块级变量（全局复用，避免重复定义） ====================
Private m_Doc         As Document       ' 当前激活文档
Private m_workPrtDoc   As PartDocument   ' 当前激活的零件文档
Private m_prt         As part           ' 当前激活的Part对象
Private m_sel         As Selection      ' 选择集对象
  '  If Not KCL.existWkPrt(m_Doc, m_workPrtDoc, m_prt, m_sel) Then Exit Sub
Private Const m_mdlname As String = "MDL_eleRename" ' UI引擎名称
Private Const mdlname As String = "MDL_eleRename"
Sub eleRename()
    If Not KCL.existWkPrt(m_Doc, m_workPrtDoc, m_prt, m_sel) Then Exit Sub
    Set mWD = New Cls_DynaWD   '创建窗口实例
    mWD.getUIcfgfromModDEC mdlname '获取窗口控件配置
    mWD.ShowToolbar mdlname
    Dim txt1: Set txt1 = mWD.GetCtrlbyname("txtname")
    txt1.text = ""
End Sub
Function fetchtxt()
   Dim txt1: Set txt1 = mWD.GetCtrlbyname("txtname")
   fetchtxt = txt1.text
End Function
Sub bodyReName_Click()
    If m_sel.count = 0 Then
        Set m_sel = KCL.Selectmulti("请选择需要重命名的实体Body（可多选）")
        If m_sel.count = 0 Then Exit Sub
    End If
    Dim lst As Object: Set lst = KCL.Initlst
    Dim itm As Object, itp As Object, i As Integer
    For i = 1 To m_sel.count
        Set itm = m_sel.item(i).Value
        Set itp = KCL.GetParent_Of_T(itm, "Body")
        If Not itp Is Nothing Then
            lst.Add itp
        ElseIf LCase(TypeName(itm)) = "body" Then
            lst.Add itm
        End If
    Next i
    m_sel.Clear
    Dim ct As Integer: ct = 1
    For Each itm In lst
        If itm.InBooleanOperation = False And itm.Name <> m_prt.MainBody.Name Then
            Select Case fetchtxt
            Case "": itm.Name = "Body." & ct: ct = ct + 1
            Case Else: itm.Name = fetchtxt & "." & ct: ct = ct + 1
            End Select
        End If
    Next
    MsgBox "实体重命名完成！共处理 " & ct - 1 & " 个Body。", vbInformation
End Sub
' =======：几何图形集（HybridBody）重命名 ====
Sub HBReName_Click()
    Dim filter: filter = "HybridBody"
    Dim OHB As Object: Set OHB = KCL.SelectItem("请选择需要重命名元素的几何图形集", filter)
    If OHB Is Nothing Then Exit Sub
    ' 2. 按类型分类重命名
    Dim HSF As HybridShapeFactory: Set HSF = m_prt.HybridShapeFactory
    Dim oshapes As HybridShapes: Set oshapes = OHB.HybridShapes
    Dim ct As Variant: ct = Array(0, 0, 0, 0, 0, 0, 0, 0)
    Dim oWF As Object, i As Integer
    For i = 1 To oshapes.count
        Set oWF = oshapes.item(i)
        Select Case HSF.GetGeometricalFeatureType(oWF)
            Case 0: oWF.Name = "aShape_" & ct(0): ct(0) = ct(0) + 1
            Case 1: oWF.Name = "point_" & ct(1): ct(1) = ct(1) + 1
            Case 2: oWF.Name = "curve_" & ct(2): ct(2) = ct(2) + 1
            Case 3: oWF.Name = "line_" & ct(3): ct(3) = ct(3) + 1
            Case 4: oWF.Name = "circle_" & ct(4): ct(4) = ct(4) + 1
            Case 5: oWF.Name = "surface_" & ct(5): ct(5) = ct(5) + 1
            Case 6: oWF.Name = "plane_" & ct(6): ct(6) = ct(6) + 1
            Case 7: oWF.Name = "solid_" & ct(7): ct(7) = ct(7) + 1
            Case Else: oWF.Name = "SHape_" & ct(7): ct(7) = ct(7) + 1
        End Select
    Next i
    MsgBox "线框元素重命名完成！共处理 " & oshapes.count & " 个元素。", vbInformation
End Sub
Sub hybodyReName_Click()
    If m_sel.count = 0 Then
        Set m_sel = KCL.Selectmulti("请选择需要重命名的shape（可多选）")
        If m_sel.count = 0 Then Exit Sub
    End If
    Dim lst As Object: Set lst = KCL.Initlst
    Dim itm As Object, itp As Object, i As Integer
    For i = 1 To m_sel.count
        Set itm = m_sel.item(i).Value
            lst.Add itm
    Next i
    m_sel.Clear
    Dim ct As Integer: ct = 1
    For Each itm In lst
            Select Case fetchtxt
            Case "": itm.Name = "HBY." & ct: ct = ct + 1
            Case Else: itm.Name = fetchtxt & "." & ct: ct = ct + 1
            End Select
    Next
    MsgBox "线框元素重命名完成！共处理 " & ct - 1 & " 个Body。", vbInformation
End Sub
