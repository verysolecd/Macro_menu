Attribute VB_Name = "Drw_myframe2"
'%lb   mingcheng_assy_val    ,  XX名称  ,  90 , 40
'%lb   cailiao_part_val     ,  XX材料  ,  90 ,  25
'%lb   wuliaobianma_assy   ,  编码    ,  53 ,  36
'%lb   wuliaobianma_assy_val ,XX物料编码, 25 ,  36
'%lb   tuhao_assy  ,  图号    ,  53,  25
'%lb   tuhao_assy_val  ,        XX图号  ,  25,  25
'%lb   gongsimingcheng ,    我们公司    ,  25 ,  50
'%lb   gongsimingcheng_eng, OUR COMPANY CO.LTD, 25,46
'%lb   tuyangbiaoji , 图样标记 ,  50 ,  15
'%lb   tuyangbiaoji_assy_val   ,  xx图样  ,  50,  8
'%lb   zhongliang_assy ,  重量    ,  30 ,  15
'%lb   zhongliang_assy_val ,  XX重量  ,  30 ,  8
'%lb   bili, 比例,10,15
'%lb   bili_assy_val ,XX比例, 10,8

'%lb   gongxxzhang ,  共X页   ,  45 ,  0
'%lb   dixxzhang   ,  第X页   ,  15 , 0

'%lb   pizhun  ,  批     准    ,  170 ,  0
'%lb   shenhe  ,  审     核    ,  170 ,  6
'%lb   biaozhunhua , 标准化,     172 ,  12
'%lb   gongyi  ,  工     艺    ,  170,  18
'%lb   jiaohe  ,  校     核    ,  170 ,  24
'%lb   sheji   ,  设     计    ,  170 ,  30
'%lb   biaoji  ,  标记    ,  174 ,  36
'%lb   chushu  ,  处数    ,  166 ,  36
'%lb genggaiwenjianhao,更改号,  156 , 36
'%lb   qianzi  ,  签字    ,  138 ,  36
'%lb   riqi    ,  日期    ,  126 ,  36


' %UI Button btn_create 创建图框
' %UI Button btn_delete 删除图框
' %UI Button btn_resize 更改图框尺寸
' %UI Button btn_update 更新图框


'该宏需要多少前缀类型？

'frame_border_
'frame_title_block_
'title_block_std_
'titleblock_text_
'titleblock_Line_
'titleblock_revision_Line_



Public ActiveDoc, Fact, Selection
Private Sheets, Sheet, targetSheet, Views, View, Texts, Text
Private Name, m_MacroID, m_DisplayFormat As String
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
      Err.Clear: Name = "none"
    End If
  On Error GoTo 0
        Dim frmDic: Set frmDic = KCL.getFrmDic ' oFrm.Res
      Select Case frmDic("btn_clicked")
      
        Case "btn_create": If (Name = "none") Then CATDrw_Creation targetSheet
        Case "btn_resize"
            If (Name <> "none") Then
             CATDrw_Resizing targetSheet
              CATDrw_Update targetSheet
            End If
          Case "btn_delete": If (Name <> "none") Then CATDrw_Deletion targetSheet
         Case Else
            MsgBox "未点击任何按钮，或按钮名称未匹配", vbExclamation
    End Select
    CATExit targetSheet
End Sub
Sub CATCreateTitleBlockFrame()

m_MacroID = "My Drawing frame"
m_NbOfRevision = 9
m_RevRowHeight = 5
m_checkRowHeight = 6
m_RulerLength = 200
s0X = Array(-180, -164, -148, -134, -120, -60)
s0Y = Array(0, 6, 36, 61)

s1X = Array(-180, -172)
s1Y = Array(0, 36, 61)

s2X = Array(-60, -50, -40, -30, -20)
s2Y = Array(0, 6, 15, 20, 46, 46, 61)

Nb_check = 6
Nb_rv = 5
'此处分为2个大区域，TitleBlock和RevisionBlock
Dim vleft, vtop
vleft = -180
vtop = 61

 '上下
    newLineH vleft + X0, X0, Y0, "TitleBlock_Frame_line_Bottom"
    newLineH vleft + X0, X0, vtop + Y0, "TitleBlock_Frame_line_Top"
 '左右
    newLineV vleft + X0, Y0, vtop + Y0, "TitleBlock_Frame_line_Left"
    newLineV X0, Y0, vtop + Y0, "TitleBlock_Frame_line_Right"

    tb_X = Array(0, -120, -60, -50, -40, -30, -20)
    tb_Y = Array(0, 6, 15, 20, 36, 46, 61)
    
    '大section 行  titleblock_frame_line_
    
    For i = 1 To UBound(tb_Y)
        Select Case i
            Case 1, 4
            newLineH tb_X(1) + X0, X0, tb_Y(i) + Y0, "TitleBlock_Frame_line_Row_" & i
            Case Else
            newLineH tb_X(2) + X0, X0, tb_Y(i) + Y0, "TitleBlock_Frame_line_Row_" & i
        End Select
    Next
  
    '大section 列
    
    newLineV tb_X(1) + X0, Y0, tb_Y(6) + Y0, "TitleBlock_Frame_line_Col_1"
    newLineV tb_X(2) + X0, Y0, tb_Y(6) + Y0, "TitleBlock_Frame_line_Col_2"
    newLineV tb_X(3) + X0, tb_Y(3) + Y0, tb_Y(5) + Y0, "TitleBlock_Frame_line_Col_3"
    newLineV tb_X(4) + X0, tb_Y(1) + Y0, tb_Y(3) + Y0, "TitleBlock_Frame_line_Col_4"
    newLineV tb_X(5) + X0, tb_Y(0) + Y0, tb_Y(1) + Y0, "TitleBlock_Frame_line_Col_5"
    newLineV tb_X(6) + X0, tb_Y(1) + Y0, tb_Y(3) + Y0, "TitleBlock_Frame_line_Col_6"

   
    
     'REV区域 行
    Rev_x = Array(-180, -164, -148, -134)
    
        For i = 2 To Nb_check 'REV区域 行
            newLineH vleft + X0, tb_X(1) + X0, Y0 + m_checkRowHeight * i, "RevisionBlock_Line_row" & i
        Next
        
        For i = 1 To UBound(Rev_x) 'REV区域 列
            newLineV Rev_x(i) + X0, Y0, tb_Y(6) + Y0, "RevisionBlock_Line_Col_1"
        Next
        
    'Rev区域 行
        For i = 1 To Nb_rv - 1
           newLineH vleft + X0, tb_X(1) + X0, Y0 + tb_Y(4) + m_RevRowHeight * i, "RevisionBlock_Line_Row" & i
        Next
     'Rev区域 列
        newLineV -172 + X0, 36 + Y0, tb_Y(6) + Y0, "RevisionBlock_Line_Col_1"



End Sub

Sub CATDrw_Creation(targetSheet)
  If Not CATInit() Then Exit Sub
  If CATCheckRef(1) Then Exit Sub 'To check whether a FTB exists already in the sheet
  CATCreateReference          'To place on the drawing a reference point
  CATFrame      'To draw the frame
  CATCreateTitleBlockFrame    'To draw the geometry
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

Sub CATDrw_Update(targetSheet)
  If Not CATInit() Then Exit Sub
  If CATCheckRef(0) Then Exit Sub
  CATDeleteTitleBlockStandard
  CATCreateTitleBlockStandard
  CATLinks
  CATColorGeometry
  CATExit targetSheet
End Sub



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
  With Sheet
  Dim Sz
  Sz = .PaperSize
  ' 0: catPaperPortrait, ' 1: catPaperLandscape,' 2: catPaperBestFit
  '2 catPaperA0, 3 catPaperA1, 4 catPaperA2, 5 catPaperA3, 6 catPaperA4,
      If .Orientation = 0 And (.PaperSize = 2 Or .PaperSize = 4 Or .PaperSize = 6) Or _
          .Orientation = 1 And (.PaperSize = 3 Or .PaperSize = 5) Then
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
    newLineV X0, Y0, m_Height - m_Offset, "Frame_Border_Left"
    newLineV Y0, m_Height - m_Offset, Y0, "Frame_Border_Right"
    If Err.Number <> 0 Then Err.Clear
  On Error GoTo 0
End Sub


Sub CATFrameText(Nb_CM_H, Nb_CM_V, Ruler, Cst_1, Cst_2)
Dim i, t
  On Error Resume Next
    'For i = Nb_CM_H To (Ruler / 2 / Cst_1 + 1) Step -1
    For i = Nb_CM_H To 1 Step -1
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
  Dim i, j
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
    Set axis1 = newLineH(X(1), X(5), Y(1), "TitleBlock_std_Line_Axis_1")
    Set axis2 = newLineV(X(4), Y(7), Y(4), "TitleBlock_std_Line_Axis_2")
    
    Selection.Clear
      SelectAll "CATDrwSearch.2DGeometry.Name=TitleBlock_std_Line_Axis_*"
      Selection.VisProperties.SetRealWidth 1, 1
      Selection.VisProperties.SetRealLineType 4, 1
    Selection.Clear
    
    newLineV X(2), Y(5), Y(2), "TitleBlock_std_Line_1"
    newLineV X(3), Y(3), Y(6), "TitleBlock_std_Line_3"
    
    CreateLine X(2), Y(2), X(3), Y(3), "TitleBlock_std_Line_2"
    CreateLine X(3), Y(6), X(2), Y(5), "TitleBlock_std_Line_4"
    Dim oCircle
    Set oCircle = Fact.CreateClosedCircle(X(4), Y(1), R1)
    oCircle.Name = "TitleBlock_std_Line_Circle_1"
    Set oCircle = Fact.CreateClosedCircle(X(4), Y(1), R2)
    oCircle.Name = "TitleBlock_std_Line_Circle_2"
    If Err.Number <> 0 Then Err.Clear
  On Error GoTo 0
End Sub

Function newLineH(iX1, iX2, iY2, iName) As Curve2D
Dim oLine, Point
  Set oLine = Fact.CreateLine(iX1, iY2, iX2, iY2)
  oLine.Name = iName
  Set Point = oLine.StartPoint 'Create the start point
  Point.Name = iName & "_start"
  Set Point = oLine.EndPoint 'Create the start point
  Point.Name = iName & "_end"
  Set newLineH = oLine
End Function

Function newLineV(iX1, iY1, iY2, iName) As Curve2D
Dim oLine, Point
  Set oLine = Fact.CreateLine(iX1, iY1, iX1, iY2)
  oLine.Name = iName
  Set Point = oLine.StartPoint 'Create the start point
  Point.Name = iName & "_start"
  Set Point = oLine.EndPoint 'Create the start point
  Point.Name = iName & "_end"
End Function

Sub CATTitleBlockText()
 Dim dec, lst
   dec = getDecCode()
   Set lst = ParseDec(dec)
     
     For Each ttx In lst
            oTxtName = "TitleBlock_Text_" & ttx("name")
            CreateTextAF ttx("val"), X0 - ttx("X"), ttx("Y") + Y0, oTxtName, catBottomCenter, 2.5  'catMiddleCenter
        Next
'    Select Case GetContext():
'      Case "LAY": Text.InsertVariable 0, 0, ActiveDoc.part.getItem("CATLayoutRoot").Parameters.item(ActiveDoc.part.getItem("CATLayoutRoot").Name + "\" + Sheet.Name + "\ViewMakeUp2DL.1\Scale")
'      Case "DRW": Text.InsertVariable 0, 0, ActiveDoc.DrawingRoot.Parameters.item("Drawing\" + Sheet.Name + "\ViewMakeUp.1\Scale")
'      Case Else: Text.Text = "XX"
'    End Select
  CATLinks
End Sub

Sub ComputeTitleBlockTranslation(ary)
  ary(0) = 0
  ary(1) = 0
  On Error Resume Next
    Set Text = Texts.getItem("Reference_" + m_MacroID) 'Get the reference text
    If Err.Number <> 0 Then
      Err.Clear
    Else
      ary(0) = X0 - Text.X
      ary(1) = Y0 - Text.Y
      Text.X = Text.X + ary(0)
      Text.Y = Text.Y + ary(1)
    End If
  On Error GoTo 0
End Sub
Sub ComputeRevisionBlockTranslation(ary)
  ary(0) = 0
  ary(1) = 0
  On Error Resume Next
    Set Text = Texts.getItem("RevisionBlock_Text_Init") 'Get the reference text
    If Err.Number <> 0 Then
      Err.Clear
    Else
      ary(0) = X0 + m_ColRev(4) - Text.X
      ary(1) = m_Height - m_Offset - 0.5 * m_RevRowHeight - Text.Y
    End If
  On Error GoTo 0
End Sub
Sub CATRemoveFrame()
  DeleteAll "CATDrwSearch.DrwText.Name=Frame_Text_*"
  DeleteAll "CATDrwSearch.2DGeometry.Name=Frame_*"
  DeleteAll "CATDrwSearch.2DPoint.Name=TitleBlock_Frame_line_*"
End Sub
Sub CATDeleteTitleBlockStandard()
  DeleteAll "CATDrwSearch.2DGeometry.Name=TitleBlock_std_Line_*"
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
    Dim txtItem
    txtItem = "TitleBlock_Text_" & wuliaobianma_assy_val
  
    Texts.getItem(txtItem).Text = ProductDrawn.PartNumber
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
        
        oPagetext = "TitleBlock_Text_" & "gongxxzhang"
        
        itSheet.Views.item(2).Texts.getItem(oPagetext).Text = "共" & CStr(nbSheet) & "页"
         oPagetext = "TitleBlock_Text_" & "dixxzhang"
        itSheet.Views.item(2).Texts.getItem(oPagetext).Text = "第" & CStr(curSheet) & "页"
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
Dim Point
  Set CreateLine = Fact.CreateLine(iX1, iY1, iX2, iY2)
  CreateLine.Name = iName
  Set Point = CreateLine.StartPoint 'Create the start point
  Point.Name = iName & "_start"
  Set Point = CreateLine.EndPoint 'Create the start point
  Point.Name = iName & "_end"
End Function
Function CreateText(iValue, iX, iY, iName)
  Set CreateText = Texts.Add(iValue, iX, iY)
  CreateText.Name = iName
  CreateText.AnchorPosition = catMiddleCenter
End Function
Function CreateTextAF(iValue, iX, iY, iName, iAnchorPosition, iFontSize)
  Set CreateTextAF = Texts.Add(iValue, iX, iY)
  CreateTextAF.Name = iName
  CreateTextAF.AnchorPosition = iAnchorPosition
  CreateTextAF.SetFontSize 0, 0, iFontSize
  CreateTextAF.TextProperties.Blanking = 0  'catBlankingInactive,catBlankingActive,  catBlankingOnGeom
End Function
Sub CATColorGeometry()
  If Not IsEmpty(CATIA) Then
    Select Case GetContext():
      'Case "DRW":
      '    SelectAll "CATDrwSearch.2DGeometry"
      '    Selection.VisProperties.SetRealColor 0,0,0,0
      '    Selection.Clear
      
        Case "DRW":
          SelectAll "CATDrwSearch.2DGeometry.Name=Frame_centerMark_*"
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



Private Function ParseDec(ByVal code As String) As Object
    Dim regEx As Object
    Dim matches As Object
    Dim match As Object
    Dim Cls_property As Object ' Scripting.Dictionary
    Set regEx = CreateObject("VBScript.RegExp")
    With regEx
        .Global = True
        .MultiLine = True
        ' 格式: ' %lb <ControlType> <ControlName> <Caption/Text>
        .Pattern = "^\s*'\s*%lb\s+\s*(\w+)\s*,\s*(.*)\s*,\s*(\d+(?:\.\d+)?)\s*,\s*(\d+(?:\.\d+)?)$"
    End With
    Dim lst, mdic
    Set lst = InitLst

    If regEx.TEST(code) Then
        Set matches = regEx.Execute(code)
        For Each match In matches
                Set mdic = InitDic
                mdic.Add "name", match.SubMatches(0)
                mdic.Add "val", match.SubMatches(1)
                mdic.Add "X", CLng(match.SubMatches(2))
                mdic.Add "Y", CLng(match.SubMatches(3))
           lst.Add mdic   'lst.Add mdic("name"), mdic
        Next
    End If
    Set ParseDec = lst
End Function

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
Sub CATDeleteTitleBlockFrame()
    DeleteAll "CATDrwSearch.2DGeometry.Name=TitleBlock_Frame_line_*"
End Sub

Sub CATDeleteRevisionBlockFrame()
    DeleteAll "CATDrwSearch.2DGeometry.Name=RevisionBlock_Line_*"
End Sub




Sub revertCST()
 With Sheet
  Sz = .PaperSize
  ' 0: catPaperPortrait, ' 1: catPaperLandscape,' 2: catPaperBestFit
  '2 catPaperA0, 3 catPaperA1, 4 catPaperA2, 5 catPaperA3, 6 catPaperA4
      If .Orientation = 0 And (.PaperSize = 2 Or .PaperSize = 4 Or .PaperSize = 6) Or _
          .Orientation = 1 And (.PaperSize = 3 Or .PaperSize = 5) Then
        Cst_1 = 52.5
        Cst_2 = 74.2
      End If
  End With

End Sub
Sub CATCreateReference()
  Set Text = Texts.Add("", X0, Y0)
  Text.Name = "Reference_" + m_MacroID
End Sub


Private Function getDecCode()
    Dim COMObjectName$     ' 获取VBA版本对应的COM对象名称
    #If VBA7 Then
        COMObjectName = "MSAPC.Apc.7.1"
    #ElseIf VBA6 Then
        COMObjectName = "MSAPC.Apc.6.2"
    #Else
        MsgBox "不支持当前VBA版本", vbExclamation + vbOKOnly
        Exit Function
    #End If
    Dim Apc As Object: Set Apc = Nothing   ' 获取APC对象
    On Error Resume Next
        Set Apc = CreateObject(COMObjectName)
        Dim ExecPjt As Object: Set ExecPjt = Apc.ExecutingProject
         Dim mdl: Set mdl = ExecPjt.VBProject.VBE.Activecodepane.codemodule
    On Error GoTo 0
    If mdl Is Nothing Then Exit Function
        Dim DecCnt
        DecCnt = mdl.CountOfDeclarationLines ' 获取声明行数
        If DecCnt < 1 Then Exit Function
        getDecCode = mdl.Lines(1, DecCnt) ' 获取声明代码
End Function


Private Function InitLst() As Object
    Set InitLst = CreateObject("System.Collections.ArrayList")
End Function


Private Function InitDic(Optional compareMode As Long = vbBinaryCompare) As Object
    Dim Dic As Object
    Set Dic = CreateObject("Scripting.Dictionary")
    Dic.compareMode = compareMode
    Set InitDic = Dic
End Function



Sub CATFrameCentringMark(Nb_CM_H, Nb_CM_V, Ruler, Cst_1, Cst_2)
   On Error Resume Next
        newLineV 0.5 * m_Width, m_Height - m_Offset, m_Height, "Frame_centerMark_Top"
        newLineV 0.5 * m_Width, Y0, 0, "Frame_centerMark_Bottom"
        
        
        newLineH 0, m_Offset, 0.5 * m_Height, "Frame_centerMark_Left"
        newLineH m_Width - m_Offset, m_Width, 0.5 * m_Height, "Frame_centerMark_Right"
    Dim i, X, Y
        For i = Nb_CM_H To Ruler / 2 / Cst_1 Step -1
          If (i * Cst_1 < 0.5 * m_Width - 1) Then
            X = 0.5 * m_Width + i * Cst_1
            CreateLine X, Y0, X, 0.25 * m_Offset, "Frame_centerMark_Bottom_" & Int(X)
            X = 0.5 * m_Width - i * Cst_1
            CreateLine X, Y0, X, 0.25 * m_Offset, "Frame_centerMark_Bottom_" & Int(X)
          End If
        Next
    
        
        For i = 1 To Nb_CM_H
          If (i * Cst_1 < 0.5 * m_Width - 1) Then
            X = 0.5 * m_Width + i * Cst_1
            CreateLine X, m_Height - m_Offset, X, m_Height - 0.25 * m_Offset, "Frame_centerMark_Top_" & Int(X)
            X = 0.5 * m_Width - i * Cst_1
            CreateLine X, m_Height - m_Offset, X, m_Height - 0.25 * m_Offset, "Frame_centerMark_Top_" & Int(X)
          End If
        Next
    
    For i = 1 To Nb_CM_V
      If (i * Cst_2 < 0.5 * m_Height - 1) Then
        Y = 0.5 * m_Height + i * Cst_2
        CreateLine Y0, Y, 0.25 * m_Offset, Y, "Frame_centerMark_Left_" & Int(Y)
        CreateLine X0, Y, m_Width - 0.25 * m_Offset, Y, "Frame_centerMark_Right_" & Int(Y)
        Y = 0.5 * m_Height - i * Cst_2
        CreateLine Y0, Y, 0.25 * m_Offset, Y, "Frame_centerMark_Left_" & Int(Y)
        CreateLine X0, Y, m_Width - 0.25 * m_Offset, Y, "Frame_centerMark_Right_" & Int(Y)
      End If
    Next
    If Err.Number <> 0 Then Err.Clear
  On Error GoTo 0
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
    CATDeleteRevisionBlockFrame
    
    CATCreateTitleBlockFrame
    CATMoveTitleBlockText TbTranslation
    
    ' Redraw revision block
   ' CATDeleteRevisionBlockFrame
'    CATCreateRevisionBlockFrame
    CATMoveRevisionBlockText RbTranslation
    ' Move the views
    CATColorGeometry
    CATMoveViews TbTranslation
    CATLinks
  End If
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

Function CATInit()
Dim osheet, GeomElems, msg, title
  CATInit = False
  On Error Resume Next
    Set osheet = Nothing
    Set osheet = CATIA.ActiveDocument.Sheets.item(1)
  On Error GoTo 0
  If osheet Is Nothing Then Exit Function
  Set Sheets = osheet.Parent
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
  If Sheet.PaperSize = catPaperA0 Or Sheet.PaperSize = catPaperA1 Or _
  (Sheet.PaperSize = catPaperUser And (m_Width > 594 Or m_Height > 594)) Then
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

Sub CAT2DL_ViewLayout(targetSheet)
  If Not CATInit() Then Exit Sub
  On Error Resume Next
    Name = Texts.getItem("Reference_" + m_MacroID).Name
  If Err.Number <> 0 Then
    Err.Clear: Name = "none"
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
Function CATCheckRef(Mode)
Dim nbtexts, i, notfound, wholename, leftText, refText
  nbtexts = Texts.count
  i = 0
  notfound = 0
  While (notfound = 0 And i < nbtexts)
    i = i + 1
    Set Text = Texts.item(i)
    wholename = Text.Name
    leftText = Left(wholename, 10)
    If (leftText = "Reference_") Then
      notfound = 1
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
