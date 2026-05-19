Attribute VB_Name = "MDL_BdyDel"
'%Title 现在要删除实体,那我问你?
'------控件清单--------------------------------------------------
'控件格式为 %UI <ControlType> <ControlName> <Caption/Text>
' %UI Label lbL_jpzcs  键盘造车手出品
' %UI CheckBox chk_vol  启用按体积筛选
' %UI CheckBox  chk_area  启用按面积筛选
' %UI TextBox   txt_value  若启用筛选，请输入值(#0.001#)
' %UI Button btn_COG  创建重心
' %UI Button btn_sel 选择实体
' %UI Button btn_del  删除实体
' %UI Button btncancel  取消

Private m_Doc        As Document
Private m_workPrtDoc As PartDocument
Private m_prt        As part
Private m_sel        As Selection
Private m_HSF        As HybridShapeFactory  ' 模块级共享，避免重复获取
Private m_spa
Private m_Bds
Private m_coord(2)
Private Const iv = 99999999
Private Const type_vol = "volume"
Private Const type_area = "area"

Private Const mdlname As String = "MDL_BdyDel"
Sub Main()
If Not initVar Then Exit Sub
Dim typeFilter, ifilter

Set mWD = New Cls_DynaWD
    mWD.getUIcfg mdlname
    mWD.Show
mv = iv
mv = mWD.Results("txt_value")
If mWD.Results("chk_area") Then typeFilter = type_area
If mWD.Results("chk_vol") Then typeFilter = type_vol

Select Case mWD.btnClicked
    Case "btn_COG"
       Set mlst = getmainlst
       Set picklst = getpicklst(mlst, typeFilter, mv)
       Call CreateCOGby(picklst)
    Case "btn_sel"
        Delbdysel_and_sub
    Case "btn_del"
        Set mlst = getmainlst
        Set picklst = getpicklst(mlst, typeFilter, mv)
'       Call delby(mlst, typeFilter, mv)
        
End Select

End Sub
Function getmainlst()
    Set getmainlst = KCL.Initlst
        For i = 1 To m_Bds.count
            Set itm = m_Bds.item(i)
            If itm.Shapes.count <> 0 And itm.InBooleanOperation = False Then getmainlst.Add itm
        Next
End Function
Function getpicklst(mlst, itype, ivalue)
    Set getpicklst = Nothing

    Dim pt
    Set pt = Nothing
    Set picklst = KCL.Initlst
    For Each itm In mlst
                Dim oSpa, oMeas
                Set oSpa = m_Doc.GetWorkbench("SPAWorkbench")
                Dim ref: Set ref = m_prt.CreateReferenceFromObject(itm)
                Set oMeas = oSpa.GetMeasurable(ref)
                itmvalue = 0
            Select Case itype
                Case type_vol
                    itmvalue = Round(oMeas.Volume * 1000000000, 3)
                Case type_area
                    itmvalue = Round(oMeas.Area * 1000000000, 3)
            End Select
                If itmvalue <> 0 Then
                    If Abs(itmvalue - ivalue) <= 0.003 Then picklst.Add itm
                Else
                    picklst.Add itm
                End If
    Next itm
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
    Dim oSpa, oMeasurable
    Set oSpa = m_workPrtDoc.GetWorkbench("SPAWorkbench")
    Set ref = m_prt.CreateReferenceFromObject(oBody)
    Set oMeasurable = oSpa.GetMeasurable(ref)
    oMeasurable.GetCOG Arrcord()
    GetBodyCOG = True
    Exit Function
ErrHandler:
    MsgBox "GetBodyCOG失败：" & Err.Description, vbExclamation
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
Sub Delbdysel_and_sub()
    Set m_HSF = m_prt.HybridShapeFactory
    Set m_Bds = m_prt.bodies
    Set itm = KCL.SelectItem("请选择要删除的body", "Body")
    Dim del_lst
    Set del_lst = Nothing
    Set del_lst = recurBD2lst(itm)
    delinlst del_lst
End Sub
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
            For Each itm In lst
                isel.Add itm
            Next
            isel.Delete: isel.Clear
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


