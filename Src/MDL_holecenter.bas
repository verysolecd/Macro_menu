Attribute VB_Name = "MDL_holecenter"
'{GP:4}
'{EP:Faceholecenter}
'{Caption:获取all孔中心}
'{ControlTipText: 提示选择面后后导出面上所有孔中心}
'{BackColor: }


Option Explicit
' ==================== 模块级变量（全局复用，避免重复定义） ====================
Private m_Doc         As Document       ' 当前激活文档
Private m_workPrtDoc   As PartDocument   ' 当前激活的零件文档
Private m_prt         As part           ' 当前激活的Part对象
Private m_sel         As Selection      ' 选择集对象
Private Const mdlname As String = "MDL_holecenter"
Sub Faceholecenter()
 If Not KCL.existWkPrt(m_Doc, m_workPrtDoc, m_prt, m_sel) Then Exit Sub
   Dim HSF: Set HSF = m_prt.HybridShapeFactory
'======= 选择要识别的面
    Dim Sel_surf, OHB, oref, oExtract, oFace, i, hole, oCtr
    Set Sel_surf = Nothing
    If m_sel.count = 0 Then
        Dim imsg: imsg = "选择要识别的面"
        Dim filter: filter = "Face,HybridShape"
        Set Sel_surf = KCL.SelectElement(imsg, filter)    ' ← 改用 SelectElement，拿到 SelectedElement
    ElseIf m_sel.count = 1 Then
        Debug.Print TypeName(m_sel.item(1).Value)
        Set Sel_surf = m_sel.item(1)
    Else
        MsgBox "请只选择单个面"
    End If
    If Not Sel_surf Is Nothing Then
        Set OHB = m_prt.HybridBodies.Add(): OHB.Name = "extracted points"
        Set oExtract = HSF.AddNewExtract(Sel_surf.Reference) ' ← 取 sel.Reference（BRep 自动带 GenericNaming）
        OHB.AppendHybridShape oExtract
        m_prt.Update
        Set oref = m_prt.CreateReferenceFromObject(oExtract)
        Set oFace = HSF.AddNewSurfaceDatum(oref)
            HSF.DeleteObjectForDatum oref
        Dim oBdry: Set oBdry = HSF.AddNewBoundaryOfSurface(oFace)
          OHB.AppendHybridShape oBdry
        m_prt.Update
        m_sel.Clear: m_sel.Add oBdry
            CATIA.StartCommand ("Disassemble")
            CATIA.RefreshDisplay = True
                MsgBox "请拆解窗口选择only domain后点击ok，再点击本窗口的ok"
            CATIA.RefreshDisplay = False
        m_sel.Clear
        i = 1
        Dim pt
        For Each hole In OHB.HybridShapes
            m_sel.Add hole
            If TypeOf hole Is HybridShapeCircleTritangent Then
                Set oref = m_prt.CreateReferenceFromObject(hole)
                Set oCtr = HSF.AddNewPointCenter(oref)
                OHB.AppendHybridShape oCtr
                Set oref = m_prt.CreateReferenceFromObject(oCtr)
                m_prt.Update
                Set pt = HSF.AddNewPointDatum(oref): pt.Name = "pt_" & i
                OHB.AppendHybridShape pt
                m_sel.Add oCtr
                i = i + 1
              Else
                m_sel.Add hole
            End If
        Next
                On Error Resume Next
                    m_sel.Delete: m_sel.Clear
                 On Error GoTo 0
     End If
     
     CATIA.RefreshDisplay = True
     Set m_sel = Nothing
     Set Sel_surf = Nothing
End Sub


