Attribute VB_Name = "Drw_myframe2"
Public ActiveDoc, Fact, Selection
Public Sheets, Sheet, targetSheet, Views, View, Texts, Text

' Module private variables for simplified access
Private m_MacroID, m_DisplayFormat As String
Private m_RevRowHeight, m_checkRowHeight, m_RulerLength As Double
Private m_Width, m_Height As Double
Private X0, Y0, m_Offset As Double

Private s0X, s0Y As Variant
Private s1X, s1Y As Variant
Private s2X, s2Y As Variant

Private Nb_check, Nb_rv, m_NbOfRevision As Integer
Private m_Col, m_Row, m_ColRev As Variant

Sub CATMain()
  If Not CATInit() Then Exit Sub
  On Error Resume Next
    Name = Texts.getItem("Reference_" + m_MacroID).Name
  If Err.Number <> 0 Then
    Err.Clear
    Name = "none"
  End If
  On Error GoTo 0
  If (Name = "none") Then
    CATDrw_Creation targetSheet
  Else
    CATDrw_Resizing targetSheet
    CATDrw_Update targetSheet
  End If
    CATExit targetSheet
End Sub
Sub SelectAll(iQuery As String)
  Selection.Clear
  Selection.Add (View)
  Selection.Search iQuery & ",sel"
End Sub
Sub DeleteAll(iQuery As String)
  Selection.Clear
  Selection.Add (View)
  Selection.Search iQuery & ",sel"
  If Selection.Count2 <> 0 Then Selection.Delete
End Sub
Sub CAT2DL_ViewLayout(targetSheet)
  If Not CATInit() Then Exit Sub
  On Error Resume Next
    Name = Texts.getItem("Reference_" + m_MacroID).Name
  If Err.Number <> 0 Then
    Err.Clear
    Name = "none"
  End If
  On Error GoTo 0
  If (Name = "none") Then
    CATDrw_Creation (targetSheet)
  Else
    CATDrw_Resizing (targetSheet)
    CATDrw_Update (targetSheet)
  End If
  CATExit (targetSheet)
End Sub

Sub CATDrw_Creation(targetSheet)
  If Not CATInit() Then Exit Sub
  If CATCheckRef(1) Then Exit Sub 'To check whether a FTB exists already in the sheet
  CATCreateReference          'To place on the drawing a reference point
  CATFrame      'To draw the frame
  myTbc
'  CATCreateTitleBlockFrame    'To draw the geometry
  CATCreateTitleBlockStandard 'To draw the standard representation
  CATTitleBlockText     'To fill in the title block
  CATColorGeometry 'To change the geometry color
  CATExit targetSheet      'To save the sketch edition
End Sub
Sub CATDrw_Deletion(targetSheet)
  If Not CATInit() Then Exit Sub
  If CATCheckRef(0) Then Exit Sub
  DeleteAll "..Name=Frame_*"
  DeleteAll "..Name=TitleBlock_*"
  DeleteAll "..Name=RevisionBlock_*"
  DeleteAll "..Name=Reference_*"
  CATExit targetSheet
End Sub
Sub CATDrw_Resizing(targetSheet)
   If Not CATInit() Then Exit Sub
  If CATCheckRef(0) Then Exit Sub
  Dim TbTranslation(2)
  ComputeTitleBlockTranslation TbTranslation
  Dim RbTranslation(2)
  ComputeRevisionBlockTranslation RbTranslation
  If TbTranslation(0) <> 0 Or TbTranslation(1) <> 0 Then
    ' Redraw Sheet Frame
    DeleteAll "CATDrwSearch.DrwText.Name=Frame_Text_*"
    DeleteAll "CATDrwSearch.2DGeometry.Name=Frame_*"
    CATFrame
    ' Redraw Standard Pictorgram
    CATDeleteTitleBlockStandard
    CATCreateTitleBlockStandard
    ' Redraw Title Block Frame
    CATDeleteTitleBlockFrame
    CATCreateTitleBlockFrame
    CATMoveTitleBlockText TbTranslation
    ' Redraw revision block
    CATDeleteRevisionBlockFrame
    CATCreateRevisionBlockFrame
    CATMoveRevisionBlockText RbTranslation
    ' Move the views
    CATColorGeometry
    CATMoveViews TbTranslation
    CATLinks
  End If
  CATExit targetSheet
End Sub
Sub CATDrw_Update(targetSheet)
  If Not CATInit() Then Exit Sub
  If CATCheckRef(0) Then Exit Sub
  CATDeleteTitleBlockStandard
  CATCreateTitleBlockStandard
  CATLinks
  CATColorGeometry
  CATExit targetSheet
End Sub
Function GetContext()
  Select Case TypeName(Sheet)
    Case "DrawingSheet"
        Select Case TypeName(ActiveDoc)
          Case "DrawingDocument": GetContext = "DRW"
          Case "ProductDocument": GetContext = "SCH"
          Case Else: GetContext = "Unexpected"
        End Select
        Case "Layout2DSheet": GetContext = "LAY"
        Case Else: GetContext = "Unexpected"
  End Select
End Function
Sub CATDrw_CheckedBy(targetSheet)
  If Not CATInit() Then Exit Sub
  If CATCheckRef(0) Then Exit Sub
  CATFillField "TitleBlock_Text_Controller_1", "TitleBlock_Text_CDate_1", "checked"
  CATExit targetSheet
End Sub
Sub CATDrw_AddRevisionBlock(targetSheet)
   If Not CATInit() Then Exit Sub
  If CATCheckRef(0) Then Exit Sub
  CATAddRevisionBlockText 'To fill in the title block
  CATDeleteRevisionBlockFrame
  CATCreateRevisionBlockFrame 'To draw the geometry
  CATColorGeometry
  CATExit targetSheet
End Sub

Function CATInit()
  CATInit = False
  On Error Resume Next
    Set oSheet = Nothing
    Set oSheet = CATIA.ActiveDocument.Sheets.item(1)
  On Error GoTo 0
  If oSheet Is Nothing Then Exit Function
  Set Sheets = oSheet.Parent
  Set ActiveDoc = Sheets.Parent
  Set targetSheet = Sheets.ActiveSheet
  Set Sheet = targetSheet
  Set Views = Sheet.Views
  Set View = Views.item("Background View") 'Get the background view  Set oView = oViews.item("Background View")
    View.Activate
  Set Texts = View.Texts
  Set Fact = View.Factory2D
  Set GeomElems = View.GeometricElements
  If GetContext() = "Unexpected" Then
    msg = "The macro runs in an inappropriate environment." & Chr(13) & "The script will terminate wihtout finishing the current action."
    title = "Unexpected environement error"
    MsgBox msg, 16, title
    Exit Function
  End If
  If Not IsEmpty(CATIA) Then
    Set Selection = CATIA.ActiveDocument.Selection
    Selection.Clear
    CATIA.HSOSynchronized = False
  End If
  
  Call initVar
    Select Case TypeName(Sheet)
      Case "DrawingSheet":
          m_Width = Sheet.GetPaperWidth
          m_Height = Sheet.GetPaperHeight
      Case "Layout2DSheet":
          m_Width = Sheet.PaperWidth
          m_Height = Sheet.PaperHeight
    End Select
  If Sheet.PaperSize = catPaperA0 Or Sheet.PaperSize = catPaperA1 Or (Sheet.PaperSize = catPaperUser And (m_Width > 594 Or m_Height > 594)) Then
    m_Offset = 20
  Else
    m_Offset = 10
  End If
  X0 = m_Width - m_Offset
  Y0 = m_Offset
  m_DisplayFormat = Array("Letter", "Legal", "A0", "A1", "A2", "A3", "A4", "A", "B", "C", "D", "E", "F", "User")(Sheet.PaperSize)
  CATInit = True 'Exit without error
End Function
Sub CATExit(targetSheet)
  If Not IsEmpty(CATIA) Then
    Selection.Clear
    CATIA.HSOSynchronized = True
  End If
    View.SaveEdition
End Sub
Sub CATCreateReference()
  Set Text = Texts.Add("", X0, Y0)
  Text.Name = "Reference_" + m_MacroID
End Sub
Function CATCheckRef(Mode)
  nbTexts = Texts.count
  i = 0
  notFound = 0
  While (notFound = 0 And i < nbTexts)
    i = i + 1
    Set Text = Texts.item(i)
    WholeName = Text.Name
    leftText = Left(WholeName, 10)
    If (leftText = "Reference_") Then
      notFound = 1
      refText = "Reference_" + m_MacroID
      If (Mode = 1) Then
        MsgBox "Frame and Titleblock already created!"
        CATCheckRef = 1
        Exit Function
      ElseIf (Text.Name <> refText) Then
        MsgBox "Frame and Titleblock created using another style:" + Chr(10) + "        " + m_MacroID
        CATCheckRef = 1
        Exit Function
      Else
        CATCheckRef = 0
        Exit Function
      End If
    End If
  Wend
  If Mode = 1 Then
    CATCheckRef = 0
  Else
    MsgBox "No Frame and Titleblock!"
    CATCheckRef = 1
  End If
End Function
Function CATCheckRev()
  SelectAll "CATDrwSearch.DrwText.Name=RevisionBlock_Text_Rev_*"
  CATCheckRev = Selection.Count2
End Function
Sub CATFrame()
  Dim Cst_1     'Length (in cm) between 2 horinzontal marks
  Dim Cst_2     'Length (in cm) between 2 vertical marks
  Dim Nb_CM_H  'Number/2 of horizontal centring marks
  Dim Nb_CM_V  'Number/2 of vertical centring marks
  Dim Ruler    'Ruler length (in cm)
  CATFrameStandard Nb_CM_H, Nb_CM_V, Ruler, Cst_1, Cst_2
  CATFrameBorder
  CATFrameCentringMark Nb_CM_H, Nb_CM_V, Ruler, Cst_1, Cst_2
  CATFrameText Nb_CM_H, Nb_CM_V, Ruler, Cst_1, Cst_2
  CATFrameRuler Ruler, Cst_1
End Sub
Sub CATFrameStandard(Nb_CM_H, Nb_CM_V, Ruler, Cst_1, Cst_2)
   Cst_1 = 74.2 '297, 594, 1189 are multiples of 74.2
  Cst_2 = 52.5 '210, 420, 841  are multiples of 52.2
' 0: catPaperPortrait, ' 1: catPaperLandscape,' 2: catPaperBestFit
  With Sheet
      If .Orientation = 0 And (.PaperSize = catPaperA0 Or .PaperSize = catPaperA2 Or .PaperSize = catPaperA4) Or _
          .Orientation = 1 And (.PaperSize = catPaperA1 Or .PaperSize = catPaperA3) Then
        Cst_1 = 52.5
        Cst_2 = 74.2
      End If
  End With
  Nb_CM_H = CInt(0.5 * m_Width / Cst_1)
  Nb_CM_V = CInt(0.5 * m_Height / Cst_2)
  Ruler = CInt((Nb_CM_H - 1) * Cst_1 / 50) * 100   'here is computed the maximum ruler length
  If m_RulerLength < Ruler Then Ruler = m_RulerLength
End Sub
Sub CATFrameBorder()
   On Error Resume Next
    newLineH Y0, X0, Y0, "Frame_Border_Bottom"
    newLineH X0, Y0, m_Height - m_Offset, "Frame_Border_Top"
'    CreateLine Y0, Y0, X0, Y0, "Frame_Border_Bottom"
'    CreateLine X0, m_Height - m_Offset, Y0, m_Height - m_Offset, "Frame_Border_Top"
'    CreateLine X0, Y0, X0, m_Height - m_Offset, "Frame_Border_Left"
'    CreateLine Y0, m_Height - m_Offset, Y0, Y0, "Frame_Border_Right"
    newLineV X0, Y0, m_Height - m_Offset, "Frame_Border_Left"
    newLineV Y0, m_Height - m_Offset, Y0, "Frame_Border_Right"
    If Err.Number <> 0 Then Err.Clear
  On Error GoTo 0
End Sub
Sub CATFrameCentringMark(Nb_CM_H, Nb_CM_V, Ruler, Cst_1, Cst_2)
   On Error Resume Next
    CreateLine 0.5 * m_Width, m_Height - m_Offset, 0.5 * m_Width, m_Height, "Frame_CentringMark_Top"
    CreateLine 0.5 * m_Width, Y0, 0.5 * m_Width, 0, "Frame_CentringMark_Bottom"
    CreateLine Y0, 0.5 * m_Height, 0, 0.5 * m_Height, "Frame_CentringMark_Left"
    CreateLine m_Width - m_Offset, 0.5 * m_Height, m_Width, 0.5 * m_Height, "Frame_CentringMark_Right"
    For i = Nb_CM_H To Ruler / 2 / Cst_1 Step -1
      If (i * Cst_1 < 0.5 * m_Width - 1) Then
        X = 0.5 * m_Width + i * Cst_1
        CreateLine X, Y0, X, 0.25 * m_Offset, "Frame_CentringMark_Bottom_" & Int(X)
        X = 0.5 * m_Width - i * Cst_1
        CreateLine X, Y0, X, 0.25 * m_Offset, "Frame_CentringMark_Bottom_" & Int(X)
      End If
    Next
    For i = 1 To Nb_CM_H
      If (i * Cst_1 < 0.5 * m_Width - 1) Then
        X = 0.5 * m_Width + i * Cst_1
        CreateLine X, m_Height - m_Offset, X, m_Height - 0.25 * m_Offset, "Frame_CentringMark_Top_" & Int(X)
        X = 0.5 * m_Width - i * Cst_1
        CreateLine X, m_Height - m_Offset, X, m_Height - 0.25 * m_Offset, "Frame_CentringMark_Top_" & Int(X)
      End If
    Next
    For i = 1 To Nb_CM_V
      If (i * Cst_2 < 0.5 * m_Height - 1) Then
        Y = 0.5 * m_Height + i * Cst_2
        CreateLine Y0, Y, 0.25 * m_Offset, Y, "Frame_CentringMark_Left_" & Int(Y)
        CreateLine X0, Y, m_Width - 0.25 * m_Offset, Y, "Frame_CentringMark_Right_" & Int(Y)
        Y = 0.5 * m_Height - i * Cst_2
        CreateLine Y0, Y, 0.25 * m_Offset, Y, "Frame_CentringMark_Left_" & Int(Y)
        CreateLine X0, Y, m_Width - 0.25 * m_Offset, Y, "Frame_CentringMark_Right_" & Int(Y)
      End If
    Next
    If Err.Number <> 0 Then Err.Clear
  On Error GoTo 0
End Sub
Sub CATFrameText(Nb_CM_H, Nb_CM_V, Ruler, Cst_1, Cst_2)
  On Error Resume Next
    For i = Nb_CM_H To (Ruler / 2 / Cst_1 + 1) Step -1
      CreateText Chr(65 + Nb_CM_H - i), 0.5 * m_Width + (i - 0.5) * Cst_1, 0.5 * m_Offset, "Frame_Text_Bottom_1_" & Chr(65 + Nb_CM_H - i)
      CreateText Chr(64 + Nb_CM_H + i), 0.5 * m_Width - (i - 0.5) * Cst_1, 0.5 * m_Offset, "Frame_Text_Bottom_2_" & Chr(65 + Nb_CM_H + i)
    Next
    For i = 1 To Nb_CM_H
      t = Chr(65 + Nb_CM_H - i)
      CreateText(t, 0.5 * m_Width + (i - 0.5) * Cst_1, m_Height - 0.5 * m_Offset, "Frame_Text_Top_1_" & t).Angle = -90
      t = Chr(64 + Nb_CM_H + i)
      CreateText(t, 0.5 * m_Width - (i - 0.5) * Cst_1, m_Height - 0.5 * m_Offset, "Frame_Text_Top_2_" & t).Angle = -90
    Next
    For i = 1 To Nb_CM_V
      t = CStr(Nb_CM_V + i)
      CreateText t, m_Width - 0.5 * m_Offset, 0.5 * m_Height + (i - 0.5) * Cst_2, "Frame_Text_Right_1_" & t
      CreateText(t, 0.5 * m_Offset, 0.5 * m_Height + (i - 0.5) * Cst_2, "Frame_Text_Left_1_" & t).Angle = -90
      t = CStr(Nb_CM_V - i + 1)
      CreateText t, m_Width - 0.5 * m_Offset, 0.5 * m_Height - (i - 0.5) * Cst_2, "Frame_Text_Right_1_" & t
      CreateText(t, 0.5 * m_Offset, 0.5 * m_Height - (i - 0.5) * Cst_2, "Frame_Text_Left_2" & t).Angle = -90
    Next
    If Err.Number <> 0 Then Err.Clear
  On Error GoTo 0
End Sub
Sub CATFrameRuler(Ruler, Cst_1)
  'Frame_Ruler_Guide -----------------------------------------------
  'Frame_Ruler_1cm   | | | | | | | | | | | | | | | | | | | | | | | |
  'Frame_Ruler_5cm   |         |         |         |         |
  On Error Resume Next
    CreateLine 0.5 * m_Width - Ruler / 2, 0.75 * m_Offset, 0.5 * m_Width + Ruler / 2, 0.75 * m_Offset, "Frame_Ruler_Guide"
    For i = 1 To Ruler / 100
      CreateLine 0.5 * m_Width - 50 * i, Y0, 0.5 * m_Width - 50 * i, 0.5 * m_Offset, "Frame_Ruler_1_" & i
      CreateLine 0.5 * m_Width + 50 * i, Y0, 0.5 * m_Width + 50 * i, 0.5 * m_Offset, "Frame_Ruler_2_" & i
      For j = 1 To 4
        CreateLine 0.5 * m_Width - 50 * i + 10 * j, Y0, 0.5 * m_Width - 50 * i + 10 * j, 0.75 * m_Offset, "Frame_Ruler_3" & i & "_" & j
        CreateLine 0.5 * m_Width + 50 * i - 10 * j, Y0, 0.5 * m_Width + 50 * i - 10 * j, 0.75 * m_Offset, "Frame_Ruler_4" & i & "_" & j
      Next
    Next
    If Err.Number <> 0 Then Err.Clear
  On Error GoTo 0
End Sub
Sub CATDeleteTitleBlockFrame()
    DeleteAll "CATDrwSearch.2DGeometry.Name=TBline_*"
End Sub
Sub myTbc()
    CATInit
  m_MacroID = "My Drawing frame"
  m_NbOfRevision = 9
  m_RevRowHeight = 5
  m_checkRowHeight = 6
  m_RulerLength = 200
  s0X = Array(-180, -164, -148, -134, -120, -60)
  s0Y = Array(0, 6, 36, 61)
  s1X = Array(-180, -172)
  s1Y = Array(0, 36, 61)
  s2X = Array(-60, -50, -36, -30, -18)
  s2Y = Array(0, 6, 15, 20, 36, 46, 61)
    Nb_check = 6
    Nb_rv = 5
  Call myNB
  Call CATCreateTitleBlockStandard
End Sub

Sub myNB()
    '上下
    newLineH s0X(0) + X0, X0, Y0, "TBline_Bottom"
    newLineH s0X(0) + X0, X0, s0Y(3) + Y0, "TBline_Top"
    '左右
    newLineV s0X(0) + X0, Y0, s0Y(3) + Y0, "TBline_Left"
    newLineV X0, Y0, s0Y(3) + Y0, "TBline_Right"
    '大section行
    newLineH s0X(0) + X0, X0, s0Y(1) + Y0, "TBline_Row_1"
    newLineH s0X(0) + X0, X0, s0Y(2) + Y0, "TBline_Row_2"
    '大section 列
    newLineV X0 + s0X(1), Y0, Y0 + s0Y(3), "TBline_Column_1"
    newLineV X0 + s0X(2), Y0, Y0 + s0Y(3), "TBline_Column_2"
    newLineV X0 + s0X(3), Y0, Y0 + s0Y(3), "TBline_Column_3"
    newLineV X0 + s0X(4), Y0, Y0 + s0Y(3), "TBline_Column_4"
    newLineV X0 + s0X(5), Y0, Y0 + s0Y(3), "TBline_Column_5"
    'check区域 行
        For i = 2 To Nb_check - 1
            newLineH X0 + s0X(0), X0 + s0X(4), Y0 + m_checkRowHeight * i, "TBline_Check_row" & i
        Next
    'Rev区域 行
        For i = 1 To Nb_rv - 1
           newLineH X0 + s0X(0), X0 + s0X(4), Y0 + s0Y(2) + m_RevRowHeight * i, "TitleBlock_Rvline_Row" & i
        Next
        '列
        newLineV s1X(1) + X0, s1Y(1) + Y0, s1Y(2) + Y0, "TitleBlock_Rvline_Col_1"
    'cfg区域 行
        For i = 1 To UBound(s2Y)
            Select Case i: Case 2, 3, 5
               newLineH X0 + s0X(5), X0, Y0 + s2Y(i), "TBline_Cfg_row" & i
            End Select
        Next
    'cfg区域 列
        For i = 1 To UBound(s2Y)
            Select Case i
                Case 1:    newLineV X0 + s2X(i), Y0 + s2Y(3), Y0 + s2Y(4), "TBline_Cfg_col" & i
                Case 2, 4: newLineV X0 + s2X(i), Y0 + s2Y(1), Y0 + s2Y(3), "TBline_Cfg_col" & i
                Case 3:    newLineV X0 + s2X(i), Y0 + s2Y(0), Y0 + s2Y(1), "TBline_Cfg_col" & i
            End Select
        Next
End Sub
Sub CATCreateTitleBlockStandard()
  Dim R1, R2, X(5), Y(7)
  R1 = 1: R2 = 2
  
  X(1) = X0 - 110 + 2 '中心线左侧
  X(2) = X(1) + 1.5   '梯形左侧
  X(3) = X(1) + 6   '梯形右侧
  
  X(4) = X(1) + 12  ' 圆心
  X(5) = X(1) + 15   '中心线右侧
  
  Y(1) = Y0 + 3  '中心
  Y(2) = Y(1) + R1
  Y(3) = Y(1) + R2
  Y(4) = Y(1) + R2 + 0.6
  Y(5) = Y(1) - R1
  Y(6) = Y(1) - R2
  Y(7) = 2 * Y(1) - Y(4)
  If Sheet.ProjectionMethod <> catFirstAngle Then
    Xtmp = X(2)
    X(2) = X(1) + X(5) - X(3)
    X(3) = X(1) + X(5) - Xtmp
    X(4) = X(1) + X(5) - X(4)
  End If
  
  Dim axis1, axis2
  On Error Resume Next
    Set axis1 = newLineH(X(1), X(5), Y(1), "TB_std_Line_Axis_1")
    Set axis2 = newLineV(X(4), Y(7), Y(4), "TB_std_Line_Axis_2")
    
    Selection.Clear
      SelectAll "CATDrwSearch.2DGeometry.Name=TB_std_Line_Axis_*"
      Selection.VisProperties.SetRealWidth 1, 1
      Selection.VisProperties.SetRealLineType 4, 1
    Selection.Clear
    
    newLineV X(2), Y(5), Y(2), "TB_std_Line_1"
    newLineV X(3), Y(3), Y(6), "TB_std_Line_3"
    
    CreateLine X(2), Y(2), X(3), Y(3), "TB_std_Line_2"
    CreateLine X(3), Y(6), X(2), Y(5), "TB_std_Line_4"
    Dim oCircle
    Set oCircle = Fact.CreateClosedCircle(X(4), Y(1), R1)
    oCircle.Name = "TB_std_Circle_1"
    Set oCircle = Fact.CreateClosedCircle(X(4), Y(1), R2)
    oCircle.Name = "TB_std_Circle_2"
    If Err.Number <> 0 Then Err.Clear
  On Error GoTo 0
End Sub

Function newLineH(iX1, iX2, iY2, iName) As Curve2D
  Set oLine = Fact.CreateLine(iX1, iY2, iX2, iY2)
  oLine.Name = iName
  Set Point = oLine.StartPoint 'Create the start point
  Point.Name = iName & "_start"
  Set Point = oLine.EndPoint 'Create the start point
  Point.Name = iName & "_end"
  Set newLineH = oLine
End Function

Function newLineV(iX2, iY1, iY2, iName) As Curve2D
  Set oLine = Fact.CreateLine(iX2, iY1, iX2, iY2)
  oLine.Name = iName
  Set Point = oLine.StartPoint 'Create the start point
  Point.Name = iName & "_start"
  Set Point = oLine.EndPoint 'Create the start point
  Point.Name = iName & "_end"
End Function

Sub CATCreateTitleBlockFrame()
    CreateLine X0 + m_Col(1), Y0, X0, Y0, "TBline_Bottom"
    CreateLine X0 + m_Col(1), Y0, X0 + m_Col(1), Y0 + m_Row(5), "TBline_Left"
    CreateLine X0 + m_Col(1), Y0 + m_Row(5), X0, Y0 + m_Row(5), "TBline_Top"
    CreateLine X0, Y0 + m_Row(5), X0, Y0, "TBline_Right"
    CreateLine X0 + m_Col(1), Y0 + m_Row(1), X0 + m_Col(5), Y0 + m_Row(1), "TBline_Row_1"
    CreateLine X0 + m_Col(1), Y0 + m_Row(2), X0 + m_Col(5), Y0 + m_Row(2), "TBline_Row_2"
    CreateLine X0 + m_Col(1), Y0 + m_Row(3), X0 + m_Col(5), Y0 + m_Row(3), "TBline_Row_3"
    CreateLine X0 + m_Col(1), Y0 + m_Row(4), X0 + m_Col(3), Y0 + m_Row(4), "TBline_Row_4"
    For i = 1 To m_NbOfRevision - 1
      CreateLine X0 + m_Col(5), Y0 + m_Row(5) / m_NbOfRevision * i, X0, Y0 + m_Row(5) / m_NbOfRevision * i, "TBline_Row_5" & i
    Next
    CreateLine X0 + m_Col(2), Y0 + m_Row(1), X0 + m_Col(2), Y0 + m_Row(3), "TBline_Column_1"
    CreateLine X0 + m_Col(3), Y0 + m_Row(1), X0 + m_Col(3), Y0 + m_Row(5), "TBline_Column_2"
    CreateLine X0 + m_Col(4), Y0 + m_Row(1), X0 + m_Col(4), Y0 + m_Row(2), "TBline_Column_3"
    CreateLine X0 + m_Col(5), Y0, X0 + m_Col(5), Y0 + m_Row(5), "TBline_Column_4"
    CreateLine X0 + m_Col(6), Y0, X0 + m_Col(6), Y0 + m_Row(5), "TBline_Column_5"
End Sub

Sub CATTitleBlockText()
  Text_01 = "我是图纸"
  Text_02 = "SCALE"
  Text_03 = "XXX"
  Text_04 = "WEIGHT (kg)"
  Text_05 = "XXX"
  Text_06 = "DRAWING NUMBER"
  Text_07 = "SHEET"
  Text_08 = "SIZE"
  Text_09 = "USER"
  Text_10 = "XXX"                ' Paper Format
  Text_11 = "DASSAULT SYSTEMES"
  Text_12 = "CHECKED BY:"
  Text_13 = "DATE:"
  Text_14 = "DESIGNED BY:"
  If Not IsEmpty(CATIA) Then
    Text_15 = CATIA.SystemService.Environ("LOGNAME")
    If Text_15 = "" Then Text_15 = CATIA.SystemService.Environ("USERNAME")
  Else
    Set Net = CreateObject("WScript.Network")
    Text_15 = Net.UserName
  End If
  CreateTextAF Text_01, X0 + m_Col(1) + 1, Y0 + 0.5 * m_Row(1), "TitleBlock_Text_Rights", catMiddleLeft, 1.5
  CreateTextAF Text_02, X0 + m_Col(1) + 1, Y0 + m_Row(2), "TitleBlock_Text_Scale", catTopLeft, 1.5
  Set Text = CreateTextAF("", X0 + 0.5 * (m_Col(1) + m_Col(2)) - 4, Y0 + m_Row(1), "TitleBlock_Text_Scale_1", catBottomCenter, 5)
  Select Case GetContext():
    Case "LAY": Text.InsertVariable 0, 0, ActiveDoc.part.getItem("CATLayoutRoot").Parameters.item(ActiveDoc.part.getItem("CATLayoutRoot").Name + "\" + Sheet.Name + "\ViewMakeUp2DL.1\Scale")
    Case "DRW": Text.InsertVariable 0, 0, ActiveDoc.DrawingRoot.Parameters.item("Drawing\" + Sheet.Name + "\ViewMakeUp.1\Scale")
    Case Else: Text.Text = "XX"
  End Select
  CreateTextAF Text_04, X0 + m_Col(2) + 1, Y0 + m_Row(2), "TitleBlock_Text_Weight", catTopLeft, 1.5
  CreateTextAF Text_05, X0 + 0.5 * (m_Col(2) + m_Col(3)), Y0 + m_Row(1), "TitleBlock_Text_Weight_1", catBottomCenter, 5
  CreateTextAF Text_06, X0 + m_Col(3) + 1, Y0 + m_Row(2), "TitleBlock_Text_Number", catTopLeft, 1.5
  CreateTextAF Text_05, X0 + 0.5 * (m_Col(3) + m_Col(4)), Y0 + m_Row(1), "TitleBlock_Text_EnoviaV5_Effectivity", catBottomCenter, 4
  CreateTextAF Text_07, X0 + m_Col(4) + 1, Y0 + m_Row(2), "TitleBlock_Text_Sheet", catTopLeft, 1.5
  CreateTextAF Text_05, X0 + 0.5 * (m_Col(4) + m_Col(5)), Y0 + m_Row(1), "TitleBlock_Text_Sheet_1", catBottomCenter, 5
  CreateTextAF Text_08, X0 + m_Col(1) + 1, Y0 + m_Row(3), "TitleBlock_Text_Size", catTopLeft, 1.5
  If (Sheet.PaperSize = 13) Then
    CreateTextAF Text_09, X0 + 0.5 * (m_Col(1) + m_Col(2)), Y0 + m_Row(2) + 2, "TitleBlock_Text_Size_1", catBottomCenter, 5
  Else
    CreateTextAF Text_10, X0 + 0.5 * (m_Col(1) + m_Col(2)), Y0 + m_Row(2) + 2, "TitleBlock_Text_Size_1", catBottomCenter, 5
  End If
  CreateTextAF Text_11, X0 + 0.5 * (m_Col(3) + m_Col(5)), Y0 + 0.5 * (m_Row(2) + m_Row(3)), "TitleBlock_Text_Company", catMiddleCenter, 5
  CreateTextAF Text_12, X0 + m_Col(1) + 1, Y0 + m_Row(4), "TitleBlock_Text_Controller", catTopLeft, 1.5
  CreateTextAF Text_05, X0 + m_Col(2) + 2.5, Y0 + 0.5 * (m_Row(3) + m_Row(4)), "TitleBlock_Text_Controller_1", catBottomCenter, 3
  CreateTextAF Text_13, X0 + m_Col(1) + 1, Y0 + 0.5 * (m_Row(3) + m_Row(4)), "TitleBlock_Text_CDate", catTopLeft, 1.5
  CreateTextAF Text_05, X0 + m_Col(2) + 2.5, Y0 + m_Row(3), "TitleBlock_Text_CDate_1", catBottomCenter, 3
  CreateTextAF Text_14, X0 + m_Col(1) + 1, Y0 + m_Row(5), "TitleBlock_Text_Designer", catTopLeft, 1.5
  CreateTextAF Text_15, X0 + m_Col(2) + 2.5, Y0 + 0.5 * (m_Row(4) + m_Row(5)), "TitleBlock_Text_Designer_1", catBottomCenter, 3
  CreateTextAF Text_13, X0 + m_Col(1) + 1, Y0 + 0.5 * (m_Row(4) + m_Row(5)), "TitleBlock_Text_DDate", catTopLeft, 1.5
  CreateTextAF "" & Date, X0 + m_Col(2) + 2.5, Y0 + m_Row(4), "TitleBlock_Text_DDate_1", catBottomCenter, 3
  CreateTextAF Text_05, X0 + 0.5 * (m_Col(3) + m_Col(5)), Y0 + m_Row(4), "TitleBlock_Text_Title_1", catMiddleCenter, 7
  For ii = 1 To m_NbOfRevision
    iY = Y0 + (ii - 0.5) * m_Row(5) / m_NbOfRevision
    CreateTextAF Chr(64 + ii), X0 + 0.5 * (m_Col(5) + m_Col(6)), iY, "TitleBlock_Text_Modif_" + Chr(64 + ii), catMiddleCenter, 2.5
    CreateTextAF "_", X0 + 0.5 * m_Col(6), iY, "TitleBlock_Text_MDate_" + Chr(64 + ii), catMiddleCenter, 2
  Next
  CATLinks
End Sub
Sub CATDeleteRevisionBlockFrame()
    DeleteAll "CATDrwSearch.2DGeometry.Name=RevisionBlock_Line_*"
End Sub
Sub CATCreateRevisionBlockFrame()
  Revision = CATCheckRev()
  If Revision = 0 Then Exit Sub
  For ii = 0 To Revision
    iX = X0
    iY1 = m_Height - Y0 - m_RevRowHeight * ii
    iY2 = m_Height - Y0 - m_RevRowHeight * (ii + 1)
    CreateLine iX + m_ColRev(1), iY1, iX + m_ColRev(1), iY2, "RevisionBlock_Line_Column_" + Chr(64 + ii) + "_1"
    CreateLine iX + m_ColRev(2), iY1, iX + m_ColRev(2), iY2, "RevisionBlock_Line_Column_" + Chr(64 + ii) + "_2"
    CreateLine iX + m_ColRev(3), iY1, iX + m_ColRev(3), iY2, "RevisionBlock_Line_Column_" + Chr(64 + ii) + "_3"
    CreateLine iX + m_ColRev(4), iY1, iX + m_ColRev(4), iY2, "RevisionBlock_Line_Column_" + Chr(64 + ii) + "_4"
    CreateLine iX + m_ColRev(1), iY2, iX, iY2, "RevisionBlock_Line_Row_" + Chr(64 + ii)
  Next
End Sub
Sub CATAddRevisionBlockText()
  Revision = CATCheckRev() + 1
  X = X0
  Y = m_Height - Y0 - m_RevRowHeight * (Revision - 0.5)
  init = InputBox("This review has been done by:", "Reviewer's name", "XXX")
  Description = InputBox("Comment to be inserted:", "Description", "None")
  If Revision = 1 Then
    CreateTextAF "REV", X + m_ColRev(1) + 1, Y, "RevisionBlock_Text_Rev", catMiddleLeft, 5
    CreateTextAF "DATE", X + m_ColRev(2) + 1, Y, "RevisionBlock_Text_Date", catMiddleLeft, 5
    CreateTextAF "DESCRIPTION", X + m_ColRev(3) + 1, Y, "RevisionBlock_Text_Description", catMiddleLeft, 5
    CreateTextAF "INIT", X + m_ColRev(4) + 1, Y, "RevisionBlock_Text_Init", catMiddleLeft, 5
  End If
  CreateTextAF Chr(64 + Revision), X + 0.5 * (m_ColRev(1) + m_ColRev(2)), Y - m_RevRowHeight, "RevisionBlock_Text_Rev_" + Chr(64 + Revision), catMiddleCenter, 5
  CreateTextAF "" & Date, X + 0.5 * (m_ColRev(2) + m_ColRev(3)), Y - m_RevRowHeight, "RevisionBlock_Text_Date_" + Chr(64 + Revision), catMiddleCenter, 3.5
  CreateTextAF Description, X + m_ColRev(3) + 1, Y - m_RevRowHeight, "RevisionBlock_Text_Description_" + Chr(64 + Revision), catMiddleLeft, 2.5
  CreateTextAF init, X + 0.5 * m_ColRev(4), Y - m_RevRowHeight, "RevisionBlock_Text_Init_" + Chr(64 + Revision), catMiddleCenter, 5
  On Error Resume Next
    Texts.getItem("TitleBlock_Text_MDate_" + Chr(64 + Revision)).Text = "" & Date
    If Err.Number <> 0 Then Err.Clear
  On Error GoTo 0
End Sub
Sub ComputeTitleBlockTranslation(TranslationTab)
  TranslationTab(0) = 0
  TranslationTab(1) = 0
  On Error Resume Next
    Set Text = Texts.getItem("Reference_" + m_MacroID) 'Get the reference text
    If Err.Number <> 0 Then
      Err.Clear
    Else
      TranslationTab(0) = X0 - Text.X
      TranslationTab(1) = Y0 - Text.Y
      Text.X = Text.X + TranslationTab(0)
      Text.Y = Text.Y + TranslationTab(1)
    End If
  On Error GoTo 0
End Sub
Sub ComputeRevisionBlockTranslation(TranslationTab)
  TranslationTab(0) = 0
  TranslationTab(1) = 0
  On Error Resume Next
    Set Text = Texts.getItem("RevisionBlock_Text_Init") 'Get the reference text
    If Err.Number <> 0 Then
      Err.Clear
    Else
      TranslationTab(0) = X0 + m_ColRev(4) - Text.X
      TranslationTab(1) = m_Height - m_Offset - 0.5 * m_RevRowHeight - Text.Y
    End If
  On Error GoTo 0
End Sub
Sub CATRemoveFrame()
  DeleteAll "CATDrwSearch.DrwText.Name=Frame_Text_*"
  DeleteAll "CATDrwSearch.2DGeometry.Name=Frame_*"
  DeleteAll "CATDrwSearch.2DPoint.Name=TBline_*"
End Sub
Sub CATDeleteTitleBlockStandard()
  DeleteAll "CATDrwSearch.2DGeometry.Name=TitleBlock_Standard*"
End Sub
Sub CATMoveTitleBlockText(Translation)
  SelectAll "CATDrwSearch.DrwText.Name=TitleBlock_Text_*"
  count = Selection.Count2
  For ii = 1 To count
    Set Text = Selection.Item2(ii).value
    Text.X = Text.X + Translation(0)
    Text.Y = Text.Y + Translation(1)
  Next
End Sub
Sub CATMoveViews(Translation)
  For i = 3 To Views.count
    Views.item(i).UnAlignedWithReferenceView
  Next
  For i = 3 To Views.count
      Set View = Views.item(i)
      View.X = View.X + Translation(0)
      View.Y = View.Y + Translation(1)
        Dim ReferenceView:   Set ReferenceView = View.ReferenceView
        If Not (ReferenceView Is Nothing) Then
              View.AlignedWithReferenceView
        End If
  Next
End Sub
Sub CATMoveRevisionBlockText(Translation)
  SelectAll "CATDrwSearch.DrwText.Name=RevisionBlock_Text_*"
  count = Selection.Count2
  For ii = 1 To count
    Set Text = Selection.Item2(ii).value
    Text.X = Text.X + Translation(0)
    Text.Y = Text.Y + Translation(1)
  Next
End Sub
Sub CATLinks()
  On Error Resume Next
  Dim ViewDocument
  Select Case GetContext():
    Case "LAY":
      If Not IsEmpty(CATIA) Then
        Set ViewDocument = CATIA.ActiveDocument.Product
      Else
        Set ViewDocument = ViewLayoutRootProduct
      End If
    Case "DRW":
      If Views.count >= 3 Then
        Set ViewDocument = Views.item(3).GenerativeBehavior.Document
      Else
        Set ViewDocument = Nothing
      End If
    Case Else: Set ViewDocument = Nothing
  End Select

  Dim ProductDrawn: Set ProductDrawn = Nothing
  For i = 1 To 8
    If TypeName(ViewDocument) = "PartDocument" Then
      Set ProductDrawn = ViewDocument.Product
      Exit For
    End If
    If TypeName(ViewDocument) = "Product" Then
      Set ProductDrawn = ViewDocument
      Exit For
    End If
    Set ViewDocument = ViewDocument.Parent
  Next
  If Not ProductDrawn Is Nothing Then
    Texts.getItem("TitleBlock_Text_EnoviaV5_Effectivity").Text = ProductDrawn.PartNumber
    Texts.getItem("TitleBlock_Text_Title_1").Text = ProductDrawn.Definition
    Dim ProductAnalysis As Analyze
    Set ProductAnalysis = ProductDrawn.Analyze
    Texts.getItem("TitleBlock_Text_Weight_1").Text = FormatNumber(ProductAnalysis.Mass, 2)
  End If
 
  Dim textFormat: Set textFormat = Texts.getItem("TitleBlock_Text_Size_1")
  textFormat.Text = m_DisplayFormat
  If Len(m_DisplayFormat) > 4 Then
    textFormat.SetFontSize 0, 0, 3.5
  Else
    textFormat.SetFontSize 0, 0, 5
  End If
  
  Dim nbSheet, curSheet
  If Not DrwSheet.IsDetail Then
    For Each itSheet In Sheets
      If Not itSheet.IsDetail Then nbSheet = nbSheet + 1
    Next
    For Each itSheet In Sheets
      If Not itSheet.IsDetail Then
        curSheet = curSheet + 1
        itSheet.Views.item(2).Texts.getItem("TitleBlock_Text_Sheet_1").Text = CStr(curSheet) & "/" & CStr(nbSheet)
      End If
    Next
  End If
  On Error GoTo 0
End Sub
Sub CATFillField(string1 As String, string2 As String, string3 As String)
  Dim TextToFill_1, TextToFill_2 As DrawingText
  Dim Person As String
  Set TextToFill_1 = Texts.getItem(string1)
  Set TextToFill_2 = Texts.getItem(string2)
  Person = TextToFill_1.Text
  If Person = "XXX" Then Person = "John Smith"
  Person = InputBox("This Document has been " + string3 + " by:", "Controller's name", Person)
  If Person = "" Then Person = "XXX"
  TextToFill_1.Text = Person
  TextToFill_2.Text = "" & Date
End Sub

Function CreateLine(iX1, iY1, iX2, iY2, iName) As Curve2D
  Set CreateLine = Fact.CreateLine(iX1, iY1, iX2, iY2)
  CreateLine.Name = iName
  Set Point = CreateLine.StartPoint 'Create the start point
  Point.Name = iName & "_start"
  Set Point = CreateLine.EndPoint 'Create the start point
  Point.Name = iName & "_end"
End Function
Function CreateText(iCaption, iX, iY, iName)
  Set CreateText = Texts.Add(iCaption, iX, iY)
  CreateText.Name = iName
  CreateText.AnchorPosition = catMiddleCenter
End Function
Function CreateTextAF(iCaption, iX, iY, iName, iAnchorPosition, iFontSize)
  Set CreateTextAF = Texts.Add(iCaption, iX, iY)
  CreateTextAF.Name = iName
  CreateTextAF.AnchorPosition = iAnchorPosition
  CreateTextAF.SetFontSize 0, 0, iFontSize
End Function
Sub CATColorGeometry()
  If Not IsEmpty(CATIA) Then
    Select Case GetContext():
      'Case "DRW":
      '    SelectAll "CATDrwSearch.2DGeometry"
      '    Selection.VisProperties.SetRealColor 0,0,0,0
      '    Selection.Clear
      
        Case "DRW":
          SelectAll "CATDrwSearch.2DGeometry.Name=Frame_CentringMark_*"
          Selection.VisProperties.SetRealWidth 1, 1
        Selection.Clear
      
      Case "LAY":
          SelectAll "CATDrwSearch.2DGeometry"
          Selection.VisProperties.SetRealColor 255, 255, 255, 0
          Selection.Clear
      'Case "SCH":
      '    SelectAll "CATDrwSearch.2DGeometry"
      '    Selection.VisProperties.SetRealColor 0,0,0,0
      '    Selection.Clear
    End Select
  End If
End Sub

Sub initVar()
  m_MacroID = "My Drawing frame"
  m_NbOfRevision = 9
  m_RevRowHeight = 10
  m_RulerLength = 200
  m_Col = Array(0, -190, -170, -145, -45, -25, -20)
  m_Row = Array(0, 4, 17, 30, 45, 60)
  m_ColRev = Array(0, -190, -175, -140, -20)
End Sub
