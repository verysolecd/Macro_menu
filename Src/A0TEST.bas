Attribute VB_Name = "A0TEST"
Sub CATIASearchExample()
  Set m_prt = CATIA.ActiveDocument.part
  Dim HSF:  Set HSF = m_prt.HybridShapeFactory
  CATIA.HSOSynchronized = False

'Set m_sel = KCL.SelectQuery(".'Volume geometry'& Color!=Yellow,.Surface& Color!=Yellow& Type!=Plane")
'
  'Set m_sel = KCL.SelectQuery(".'Volume geometry'+ .Surface& Color!=Yellow& Type!=Plane")
Set m_sel = KCL.SelectQuery("CatPrtSearch.Curve-CatPrtSearch.Line")

'  Set m_sel = CATIA.ActiveDocument.Selection
'  m_sel.Search ("CatPrtSearch.Curve-CatPrtSearch.Line,all")
  
    Color = Array(255, 255, 0)
    m_sel.VisProperties.SetRealColor Color(0), Color(1), Color(2), 0 '(R, G, B, Inheritance=1)
    m_sel.Clear
CATIA.HSOSynchronized = True
End Sub
Function getQuerylst(iQry, Optional ByVal rng As Variant = Nothing)
    Call KCL.SelectQuery(iQry, rng)
    Set getQuerylst = KCL.Initlst
    For i = 1 To CATIA.ActiveDocument.Selection.count
        Dim shp: Set shp = CATIA.ActiveDocument.Selection.item(i).Value
        getQuerylst.Add shp
    Next
    CATIA.ActiveDocument.Selection.Clear
End Function

