Attribute VB_Name = "MDL_BdyDel"

'Attribute VB_Name = "MDL_BdyDel"
'{GP:4}
'{Ep:delwithBody}
'{Caption:实体处理}
'{ControlTipText:各种实体处理}
'{BackColor: }

'%Title 实体操作配置?
'------控件清单--------------------------------------------------
'控件格式为 %UI <ControlType> <ControlName> <Caption/Text>
' %UI Label lbL_jpzcs  键盘造车手出品
' %UI CheckBox chk_bysel  按选择筛选(同体积)
' %UI Button btn_copy 实体复制
' %UI Button btn_COG  创建重心
' %UI Button btn_delbysel  删除实体(同体积)
' %UI Button btn_delrecur 删除实体(含子级)
' %UI Button btncancel  取消

Private m_Doc        As Document
Private m_workPrtDoc As PartDocument
Private m_prt        As part
Private m_sel        As Selection
Private m_HSF  As HybridShapeFactory
Private m_spa
Private m_Bds

Private mWD As Cls_DynaWD
Private Const m_defvol = 998877.99
Private m_coord(2)
Private Const type_vol = "volume"
Private Const mdlname As String = "MDL_BdyDel"
Sub delwithBody()
    If Not initVar Then Exit Sub
    Set mWD = New Cls_DynaWD
    mWD.ShowToolbar mdlname
End Sub

'=== 辅助：实时读勾选框值（非模态下不能用 Results） ===
Private Function getBysel() As Boolean
    getBysel = False
    If mWD Is Nothing Then Exit Function
    Dim chk As Object: Set chk = mWD.GetCtrlbyname("chk_bysel")
    If Not chk Is Nothing Then getBysel = CBool(chk.Value)
End Function

'=== 辅助：统一计算筛选参数 ===
Private Sub applyFilter(ByRef typeFilter As String, ByRef selitmValue As Double)
    typeFilter = ""
    selitmValue = 0
    If getBysel() Then
        typeFilter = type_vol
        selitmValue = Round(getSelmeas.Volume * 1000000000, 4)
    End If
End Sub

'=== 独立按钮回调 ===
Sub btn_COG_click()
    If Not initVar Then Exit Sub
    Dim typeFilter As String, selitmValue As Double
    applyFilter typeFilter, selitmValue
    Dim mlst As Object, picklst As Object
    Set mlst = getBdyLst
    Set picklst = getpicklst(mlst, typeFilter, selitmValue)
    Call CreateCOGby(picklst)
End Sub

Sub btn_copy_click()
    If Not initVar Then Exit Sub
    Dim typeFilter As String, selitmValue As Double
    applyFilter typeFilter, selitmValue
    Dim mlst As Object, picklst As Object, itm As Object
    Set mlst = getBdyLst
    Set picklst = getpicklst(mlst, typeFilter, selitmValue)
    m_sel.Clear
    For Each itm In picklst
        m_sel.Add itm
    Next
    m_sel.Copy
End Sub

Sub btn_delbysel_click()
    If Not initVar Then Exit Sub
    Dim typeFilter As String, selitmValue As Double
    ' 同体积删除：强制 type_vol
    typeFilter = type_vol
    selitmValue = Round(getSelmeas.Volume * 1000000000, 4)
    Dim mlst As Object, picklst As Object
    Set mlst = getBdyLst
    Set picklst = getpicklst(mlst, typeFilter, selitmValue)
    If picklst Is Nothing Or picklst.count = 0 Then Exit Sub
    If MsgBox("将删除 " & picklst.count & " 个同体积实体，确认？", vbYesNo + vbExclamation) <> vbYes Then Exit Sub
    delinlst picklst
End Sub

Sub btn_delrecur_click()
    If Not initVar Then Exit Sub
    Dim bdysel As Object: Set bdysel = Nothing
    If m_sel.Count2 = 1 Then
        Set bdysel = m_sel.Item2(1).Value
    Else
        Set bdysel = KCL.SelectItem("请选择实体", "Body")
    End If
    If bdysel Is Nothing Then Exit Sub
    Dim del_lst As Object: Set del_lst = recurBD2lst(bdysel)
    If del_lst Is Nothing Or del_lst.count = 0 Then Exit Sub
    If MsgBox("将删除 " & del_lst.count & " 个实体(含子级)，确认？", vbYesNo + vbExclamation) <> vbYes Then Exit Sub
    delinlst del_lst
End Sub

Sub btncancel_click()
    Set mWD = Nothing
End Sub

Private Function getSelmeas()
      Set getSelmeas = Nothing
      Dim bdysel:  Set bdysel = Nothing
        If m_sel.Count2 = 1 Then
            Set bdysel = m_sel.Item2(1).Value
        Else
            Set bdysel = KCL.SelectItem("请选择实体", "Body")
        End If
        If bdysel Is Nothing Then Exit Function
    Set getSelmeas = KCL.GetMeas(bdysel)
End Function
Function getBdyLst()
    Set getBdyLst = KCL.Initlst
        For i = 1 To m_Bds.count
            Set itm = m_Bds.item(i)
            If itm.Shapes.count <> 0 And itm.InBooleanOperation = False Then getBdyLst.Add itm
        Next
End Function
Function getpicklst(mlst, Optional ByVal iType As String, Optional ByVal flValue As Double = 0)
        Set getpicklst = Nothing
        Set picklst = KCL.Initlst
    Dim omeas, itmValue, itm
    For Each itm In mlst
          If itm.Shapes.count <> 0 And itm.InBooleanOperation = False Then
                    itmValue = 0
                 Select Case iType
                    Case type_vol
                        Set omeas = KCL.GetMeas(itm)
                        itmValue = Round(omeas.Volume * 1000000000, 4)
                    Case Else
                        itmValue = 0
                  End Select
               If itmValue <> 0 Then
                        If Abs(itmValue - flValue) <= 0.0003 Then picklst.Add itm
               Else
                        picklst.Add itm
               End If
         End If
     Next
    Set getpicklst = picklst
End Function
Sub CreateCOGby(lst)
    Set m_hB = m_prt.HybridBodies.Add()
    m_hB.Name = "cog_bodies"
    For Each itm In lst
        If GetBodyCOG(itm, m_coord) Then Set pt = CreatePt(m_coord)
        If Not KCL.IsNothing(pt) Then m_hB.AppendHybridShape pt
    Next
    m_prt.Update
End Sub

' ════════════════════════════════════════════
'  单元函数 1：获取 Body 质心坐标
'  输入：oBody  — 目标 Body
'  输出：arrCOG — Double(2)，(X, Y, Z)，单位 mm
'         返回 True 表示成功
' ════════════════════════════════════════════
Private Function GetBodyCOG(ByVal oBody As body, ByRef Arrcord()) As Boolean
    GetBodyCOG = False
    On Error GoTo ErrHandler
    Set omeas = KCL.GetMeas(oBody)
    omeas.GetCOG Arrcord()
    GetBodyCOG = True
    Exit Function
ErrHandler:
    MsgBox "GetBodyCOG失败：" & Err.Description, vbExclamation: Err.Clear
End Function
' ════════════════════════════════════════════
'  单元函数 2：在指定几何集中按坐标创建点
'  输入：oHB    — 目标 HybridBody（几何集）
'         arrCOG — Double(2)，(X, Y, Z)，单位 mm
'         sName  — 点名称（可选，默认 "Point"）
'  输出：返回创建的 HybridShapePointExplicit
' ════════════════════════════════════════════

Function CreatePt(ByRef arrCOG)
 If Not CanExecute("Productdocument,PartDocument") Then Exit Function
    On Error Resume Next
        Dim oDoc: Set oDoc = CATIA.ActiveDocument
        Dim workPrtDoc: Set workPrtDoc = KCL.get_workPartDoc
        Dim oprt: Set oprt = Nothing: Set oprt = workPrtDoc.part
    Err.Clear
    On Error GoTo 0
    If IsNothing(oprt) Then: MsgBox "No activated Part": Exit Function
    Set HSF = oprt.HybridShapeFactory
    On Error GoTo ErrHandler
    Dim oPoint
    Set oPoint = HSF.AddNewPointCoord(arrCOG(0), arrCOG(1), arrCOG(2))
    Set CreatePt = oPoint
    oprt.Update
    Exit Function
ErrHandler:
    MsgBox "[" & mdlname & "] CreatePointInSet 失败：" & Err.Description, vbExclamation
    Set CreatePt = Nothing
Exit Function

End Function

Function recurBD2lst(ByVal bd As body, Optional ByRef lst = Nothing)
   If lst Is Nothing Then Set lst = KCL.Initlst
   Dim ibd: Set ibd = bd
   lst.Add ibd
    If ibd.Shapes.count <> 0 Then
        For i = 1 To ibd.Shapes.count
        On Error Resume Next
            Set itm = ibd.Shapes.item(i).body
            If LCase(TypeName(itm)) = "body" Then
                Call recurBD2lst(itm, lst)
            End If
        Next
    End If
    Set recurBD2lst = lst
End Function

Sub delinlst(lst)
        Dim isel: Set isel = CATIA.ActiveDocument.Selection: isel.Clear
        KCL.CatiaFreeze
        On Error Resume Next
            For Each itm In lst
                isel.Add itm
            Next
            isel.Delete: isel.Clear
          Err.Clear
        On Error GoTo 0
    KCL.CatiaFreeze False
End Sub
Function initVar()
    initVar = False
    On Error GoTo ErrHandler
    If Not KCL.existWkPrt(m_Doc, m_workPrtDoc, m_prt, m_sel) Then Exit Function
    Set m_HSF = m_prt.HybridShapeFactory
    Set m_Bds = m_prt.bodies
    Set m_spa = m_workPrtDoc.GetWorkbench("SPAWorkbench")
    initVar = True
    Exit Function
ErrHandler:
    MsgBox "初始化失败：" & Err.Description, vbExclamation: Error.Clear
    Exit Function
End Function



