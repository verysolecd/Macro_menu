Attribute VB_Name = "MDL_WeldColor"
'{GP:4}
'{EP:Yellow_Weld}
'{Caption:焊缝改色}
'{ControlTipText: 将所有焊缝改为黄色}
'{BackColor:12648447}
Private m_Doc         As Document       ' 当前激活文档
Private m_workPrtDoc   As PartDocument   ' 当前激活的零件文档
Private m_prt         As part           ' 当前激活的Part对象
Private m_sel         As Selection      ' 选择集对象
Private Const TYPE_SWEEP As Long = 7
Private Const mdlname As String = "MDL_WeldColor"
Sub Yellow_Weld()
   If Not KCL.existWkPrt(m_Doc, m_workPrtDoc, m_prt, m_sel) Then Exit Sub
   If m_prt Is Nothing Then Exit Sub
    Dim c, Color, i
    Dim HSF:  Set HSF = m_prt.HybridShapeFactory
    Dim sweeps: Set sweeps = KCL.Initlst
    m_sel.Clear
    CATIA.HSOSynchronized = False
'  '  sel.Search "Type=*,scr"'
'    'm_sel.Search "CATGMOSearch.Surface,all" ' Sweep 是曲面的子类型,
'    m_sel.Search ("CATGMOSearch.HybridShape,all")
  Set m_sel = KCL.SelectQuery("(.'Volume geometry'+.Surface& Type!=Plane)& Color!=Yellow")
    Color = Array(255, 255, 0)
    m_sel.VisProperties.SetRealColor Color(0), Color(1), Color(2), 0 '(R, G, B, Inheritance=1)
    m_sel.Clear
        CATIA.HSOSynchronized = True
End Sub
