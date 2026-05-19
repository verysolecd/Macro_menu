Attribute VB_Name = "A0TEST"
'%Title 现在要删除实体,那我问你?
'------控件清单--------------------------------------------------
'控件格式为 %UI <ControlType> <ControlName> <Caption/Text>
' %UI Label lbL_jpzcs  键盘造车手出品
' %UI CheckBox chk_vol  启用按体积筛选
' %UI CheckBox  chk_area  启用按面积筛选
' %UI TextBox   txt_value  若启用筛选，请输入值(#0.001#)
' %UI Button btn_COG  创建重心
' %UI Button btndel  删除实体
' %UI Button btncancel  取消

Private m_Doc        As Document
Private m_workPrtDoc As PartDocument
Private m_prt        As part
Private m_sel        As Selection
Private m_HSF        As HybridShapeFactory  ' 模块级共享，避免重复获取
Private m_spa
Private m_Bds
Private m_coord(2)

Private Const mdlname As String = "A0TEST"
Sub Main()
If Not initVar Then Exit Sub

Set mWD = New Cls_DynaWD
    mWD.getUIcfg mdlname
    mWD.Show
Select Case mWD.btnClicked
    Case "btnOK"
        '===========路径设置
        If mWD.Results("chk_path") Then
            outputpath = IIf(oDoc.path = "", "", oDoc.path)
        Else:
            outputpath = KCL.selFdl()
        End If
End Select

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
Sub CreateCOG()
    Set m_hB = m_prt.HybridBodies.Add()
    m_hB.Name = "cog_bodies"
    Dim itm, pt
    Set pt = Nothing
    For i = 1 To m_Bds.count
        Set itm = m_Bds.item(i)
        If itm.Shapes.count <> 0 And itm.InBooleanOperation = False Then
           If GetBodyCOG(itm, m_coord) Then Set pt = CreatePt(m_coord)
           If Not KCL.IsNothing(pt) Then m_hB.AppendHybridShape pt
           m_prt.Update
        End If
    Next i
End Sub

Sub Delbdysel_and_sub()
  
    Set m_HSF = m_prt.HybridShapeFactory
    Set m_Bds = m_prt.bodies
    Set itm = KCL.SelectItem("请选择要删除的body", "Body")
    Dim del_lst
    Set del_lst = Nothing
    Set del_lst = recurBD2lst(itm)
    delinlst del_lst
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
' ════════════════════════════════════════════
'  组合示例：给 MainBody 创建质心点
' ════════════════════════════════════════════
Sub CreateMainBodyCOGPoint()
    If Not KCL.existWkPrt(m_Doc, m_workPrtDoc, m_prt, m_sel) Then Exit Sub
    Set m_HSF = m_prt.HybridShapeFactory
    m_prt.Update
    
    Dim arrCOG() As Double
    If Not GetBodyCOG(m_prt.MainBody, arrCOG) Then Exit Sub
    
    ' 在根级新建几何集
    Dim oHB As HybridBody
    Set oHB = m_prt.HybridBodies.Add()
    oHB.Name = "COG"
    Dim oPoint As HybridShapePointExplicit
    Set oPoint = CreatePointInSet(oHB, arrCOG, m_prt.MainBody.Name & "_COG")
    m_prt.Update
    MsgBox "质心点已创建  X=" & Format(arrCOG(0), "0.00") & _
           "  Y=" & Format(arrCOG(1), "0.00") & _
           "  Z=" & Format(arrCOG(2), "0.00") & "  (mm)"
End Sub

Sub CATIASearchExample()
 Set oprt = CATIA.ActiveDocument.part
Set osel = CATIA.ActiveDocument.Selection
 Set colls = oprt.bodies
 Set lst = KCL.Initlst
 For Each bd In colls
    If bd.Shapes.count = 0 Then lst.Add bd
Next bd
For Each bd In lst
osel.Add bd
Next
osel.Delete
 
End Sub
Sub removebyVolume()
    If Not KCL.existWkPrt(m_Doc, m_workPrtDoc, m_prt, m_sel) Then Exit Sub
        Set m_HSF = m_prt.HybridShapeFactory
        Set m_Bds = m_prt.bodies
        
     Dim oSpa, oMeasurable, i
        Set oSpa = m_workPrtDoc.GetWorkbench("SPAWorkbench")
        
        
        vol2d = getuserInput("请输入要删除的body体积，三位小数mm3")
        If vol2d <> "" And vol2d <> 0 Then

    
    Set osel = CATIA.ActiveDocument.Selection
    osel.Clear
    Set lst = KCL.Initlst
    For i = 1 To oBodies.count
        Set itm = oBodies.item(i)
        If itm.Shapes.count <> 0 Then
            Set ref = oPart.CreateReferenceFromObject(itm)
            Set oMeasurable = oSpa.GetMeasurable(ref)
'                Debug.Print oBody.Name & "体积" & _
'                Round(oMeasurable.Volume * 1000000000, 3) & "mm3"
               If Abs(Round(oMeasurable.Volume * 1000000000, 3) - Vol) <= 0.003 Then
               Debug.Print itm.Name
               lst.Add itm
               End If
        End If
    Next i
    osel.Clear
        On Error Resume Next
            delinlst lst
        On Error GoTo 0
    End If
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

