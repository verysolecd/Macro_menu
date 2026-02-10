VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Cls_XLM"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private header_bom(1 To 17)
Private header_rv(1 To 14)
Private hdgx(1 To 15)
Private wsEVT As Cls_WsEvt
Private wb As Object
Private ws As Object

' 在类初始化时检查实例
Private Const mdlname As String = "Cls_XLM"
Private Sub Class_Initialize()
On Error Resume Next
    If gws Is Nothing Then
        Set xlAPP = CreateObject("Excel.Application")
        Set gwb = xlAPP.Workbooks.Add
        Set gws = gwb.ActiveSheet
        Set wb = gwb: Set ws = gws
        ws.name = "bom"
    Else
        Set xlAPP = GetObject("Excel.Application")
        Set wb = gwb: Set ws = gws
    End If
    
If Error.Number <> 0 Then Err.Clear
On Error GoTo 0
    Set wsEVT = New Cls_WsEvt
    wsEVT.Init wb
    iniarr
End Sub
Private Sub iniarr()
    header_bom(1) = "No." & Chr(10) & "编号"
    header_bom(2) = "Layout" & Chr(10) & "层级"
    header_bom(3) = "PN" & Chr(10) & "零件号"
    header_bom(4) = "Nomenclature" & Chr(10) & "英文名称"
    header_bom(5) = "Definition" & Chr(10) & "中文名称"
    header_bom(6) = "Picture" & Chr(10) & "图像"
    header_bom(7) = "Quantity" & Chr(10) & "数量(PCS)"
    header_bom(8) = "Weight" & Chr(10) & "单质量"
    header_bom(9) = "sum_Weight" & Chr(10) & "总质量"
    header_bom(10) = "Material" & Chr(10) & "材料"
    header_bom(11) = "Thickness" & Chr(10) & "厚度(mm)"
    header_bom(12) = ""
    header_bom(13) = "Material" & Chr(10) & "材料"
    header_bom(14) = "Density" & Chr(10) & "密度(kg/m^3)"
    header_bom(15) = "UTS" & Chr(10) & "抗拉(Mpa)"
    header_bom(16) = "YS" & Chr(10) & "屈服(Mpa)"
    header_bom(17) = "EL" & Chr(10) & "延伸率(%)"
   

    header_rv(1) = "零件号" & Chr(10) & "Partnumber"
    header_rv(2) = "更改" & Chr(10) & "件号"
    header_rv(3) = "英文名称" & Chr(10) & "Nomenclature"
    header_rv(4) = "更改" & Chr(10) & "英文名"
    header_rv(5) = "中文名称" & Chr(10) & "Definition"
    header_rv(6) = "更改" & Chr(10) & "中文名"
    header_rv(7) = "实例名" & Chr(10) & "InstanceName"
    header_rv(8) = "更改" & Chr(10) & "实例名"
    header_rv(9) = "材料" & Chr(10) & "material"
    header_rv(10) = "定义" & Chr(10) & "材料"
    header_rv(11) = "密度" & Chr(10) & "Density"
    header_rv(12) = "更改" & Chr(10) & "密度"
    header_rv(13) = "质量" & Chr(10) & "Mass"
    header_rv(14) = "厚度" & Chr(10) & "Thickness"
    
    hdgx(1) = "NO."
    hdgx(2) = "TIRE"
    hdgx(3) = ""
    hdgx(4) = ""
    hdgx(5) = ""
    hdgx(6) = "PART NO."
    hdgx(7) = "DRAWING NO."
    hdgx(8) = "中文名称"
    hdgx(9) = "ITEM NAME"
    hdgx(10) = "Ver."
    hdgx(11) = "QTY."
    hdgx(12) = "UNIT"
    hdgx(13) = "UNIT WEIGHT"
    hdgx(14) = "TOTAL WEIGHT"
    hdgx(15) = "MATERIAL"
End Sub
Private Sub set_Header(header, Optional ByVal idxrow = 2)
    Dim lastCol
    lastCol = UBound(header) - LBound(header) + 1
    With ws
    .Range(.Cells(idxrow, 1), .Cells(idxrow, lastCol)).Value = header
    .Range(.Cells(idxrow, 1), .Cells(idxrow, lastCol)).Interior.color = RGB(190, 190, 190)
    End With
End Sub
Private Sub RVbackColor()
        Dim mapcells: mapcells = Array(0, 1, 3, 5, 7, 9, 11, 13, 14)
        Dim lastrow: lastrow = ws.Cells(ws.Rows.count, 1).End(xlUp).row
        For rowNb = 2 To lastrow
            For i = 1 To UBound(mapcells)
                ws.Cells(rowNb, mapcells(i)).Interior.color = RGB(190, 190, 190)
            Next i
        Next rowNb
End Sub
     

Public Sub clear_table()
    ws.UsedRange.ClearContents
End Sub

Public Sub setxlHead(ByVal S$)
Select Case S
    Case "bom": Call set_Header(header_bom, 1)
    Case "rv": Call set_Header(header_rv, 1): RVbackColor
    Case "gxbom": Call set_Header(hdgx, 4)
 End Select
End Sub

Public Sub freesheet()
Set ws = Nothing
Set gws = Nothing
Set gwb = Nothing
Set xlAPP = Nothing
End Sub

Public Function extract_ary()
 Dim resultAry As Variant
 Dim cntx, cnty
    cntx = ws.Cells(ws.Rows.count, 1).End(xlUp).row
    cnty = 14
 ReDim resultAry(1 To cntx, 1 To cnty)
  For idrow = 2 To cntx
         Dim tempArr As Variant
            tempArr = ws.Rows(idrow).Resize(1, cnty).Value
        For j = 1 To cnty
            resultAry(idrow, j) = tempArr(1, j)
        Next j
   Next idrow
    extract_ary = resultAry ' 返回提取的数据
End Function

Private Sub LvMg()
xlhide
    Dim row_num As Long ' 声明 row_num 变量
    Dim cell ' 声明 cell 变量
     With ws
            .Cells.ClearOutline
            .Outline.AutomaticStyles = False
            .Outline.SummaryRow = xlAbove
            .Outline.SummaryColumn = xlRight
            Set LVcol = .Columns(2)
            LVcol.HorizontalAlignment = xlLeft
    End With
    For row_num = 2 To ws.Cells(ws.Rows.count, 2).End(xlUp).row ' 使用 xlApp.xlUp
        Dim cell_value As Variant
        cell_value = ws.Cells(row_num, 2).Value
        If Not IsEmpty(cell_value) Then
            ws.Rows(row_num).OutlineLevel = cell_value
        End If
    Next
    ' 设置对齐方式和缩进级别
    For Each cell In ws.Range("B2:B" & ws.Cells(ws.Rows.count, 2).End(xlUp).row) ' 使用 xlApp.xlUp
        If Not IsEmpty(cell) Then
            cell.HorizontalAlignment = xlLeft ' 使用 xlApp.xlLeft
            If IsNumeric(cell.Value) Then
                cell.IndentLevel = cell.Value ' 修改缩进级别与单元格值相等
            End If
        End If
    Next
'   lastRow = ws.Cells(Rows.count, 3).End(xlUp).Row
'   ws.Rows(lastRow + 3 & ":" & Rows.count).EntireRow.Hidden = True
'   Columns("Q:Q").Select
'   Range(Selection, Selection.End(xlToRight)).Select
'   Selection.EntireColumn.Hidden = True
End Sub
' -------------------------------------------------------------------------
' Typed Injection Methods (Bomline)
' -------------------------------------------------------------------------



' 通用数组注入方法 (支持 legacy/debug 及 RW_Revise 的调用)
' RW_Revise 调用格式: inject_ary(data, startRow, targetCols, sourceIndices)
' Debug 调用格式: inject_ary(data)
Public Sub inject_ary(ByVal data As Variant, Optional ByVal startrow As Long = 2, Optional ByVal targetCols As Variant, Optional ByVal sourceIndices As Variant)
    xlhide
    ' 如果只传了1个参数（兼容旧/调试调用），直接写入
    If IsMissing(targetCols) Or IsMissing(sourceIndices) Then
        Dim R As Long, C As Long
        On Error Resume Next
        R = UBound(data, 1)
        C = UBound(data, 2)
        If Err.Number <> 0 Then Exit Sub
        On Error GoTo 0
        
        Dim lastRowSimple As Long
        lastRowSimple = startrow + R - 1
        With ws
            .Range(.Cells(startrow, 1), .Cells(lastRowSimple, C)).Value = data
        End With
        Call LvMg
        ws.name = "BOM_" & Format(Now, "yymmdd_hhmm")
        xlshow
        Exit Sub
    End If
    ' 4参数模式 (RW_Revise 用法)
    Dim rCount As Long, cCount As Long
    rCount = UBound(data, 1)
    cCount = UBound(targetCols) ' RW_Revise 传入的是 Array(...)，通常从0开始
    Dim i As Long, j As Long, tCol As Long, sIdx As Long
    Dim val As Variant
    On Error Resume Next
    ' Loop through data rows (attLv2Prd返回是1-based)
    For i = 1 To rCount
        ' Loop through mapping columns
        For j = LBound(targetCols) To UBound(targetCols)
            tCol = targetCols(j)
            sIdx = sourceIndices(j)
            
            ' RW_Revise 中 targetCols 可能包含 0 (占位符)
            If tCol > 0 Then
                ' sIdx 是数据源的列索引 (需确保在范围内)
                If sIdx > 0 And sIdx <= UBound(data, 2) Then
                    val = data(i, sIdx)
                    ws.Cells(startrow + i - 1, tCol).Value = val
                End If
            End If
        Next j
    Next i
    On Error GoTo 0
    
    ws.name = "Revise_" & Format(Now, "hhmm")
    ' RW_Revise calls xlshow/visible explicitly after this
End Sub

Public Sub inject_Bom(arr2D As Variant, Optional ByVal startrow As Long = 2)
xlhide:
    Dim rowCount As Long: rowCount = UBound(arr2D, 1)
    Dim colCount As Long: colCount = UBound(arr2D, 2)
    lastrow = rowCount + startrow - 1
    Call setxlHead("bom")
    With ws
        .Range(.Cells(startrow, 1), .Cells(lastrow, colCount)).Value = arr2D
        If lastrow >= 2 Then
            With .Range(.Cells(startrow, 9), .Cells(lastrow, 9))
                .Formula = "=RC[-2]*RC[-1]"
                .NumberFormat = "0.000"
            End With
        End If
    End With
    Call LvMg
    ws.name = "BOM_" & Format(Now, "yymmdd_hhmm")
End Sub
' 转换函数：Bomline() → 二维数组 (标准BOM格式)

Public Sub inject_GXbom(arr2D As Variant, Optional ByVal startrow As Long = 5)
    xlhide
    Dim rowCount As Long: rowCount = UBound(arr2D, 1)
    Dim colCount As Long: colCount = UBound(arr2D, 2)
    lastrow = rowCount + startrow - 1
    Call set_Header(hdgx, startrow - 1)
    With ws
        .Range(.Cells(startrow, 1), .Cells(lastrow, colCount)).Value = arr2D
    End With
    ws.name = "GX_" & Format(Now, "yymmdd_hhmm")
End Sub

Public Sub inject_RvData(data() As Bomline, Optional ByVal startrow As Long = 2)
    xlhide
    Dim cnt As Long: cnt = UBound(data)
    Dim targetCols: targetCols = Array(1, 3, 5, 7, 9, 11, 13, 14)
    Dim buffer() As Variant: ReDim buffer(1 To cnt, 1 To 1)
    Dim C As Long, i As Long
    startrow = 2
    Call setxlHead("rv")
        For C = 0 To UBound(targetCols)
        For i = 1 To cnt
            With data(i)
                Select Case C
                    Case 0: buffer(i, 1) = .partNumber
                    Case 1: buffer(i, 1) = .Nomenclature
                    Case 2: buffer(i, 1) = .Definition
                    Case 3: buffer(i, 1) = .InstanceName
                    Case 4: buffer(i, 1) = .Material
                    Case 5: buffer(i, 1) = .Density
                    Case 6: buffer(i, 1) = .Mass
                    Case 7: buffer(i, 1) = .Thickness
                End Select
            End With
        Next i
        ws.Cells(startrow, targetCols(C)).Resize(cnt, 1).Value = buffer
    Next C
        Call RVbackColor
    xlshow
End Sub

Sub inject_pic(irow, pncol, Piccol, ipath)
        Dim i&, rng As Range
        Dim iPic As Shape
        Dim cntrow
        ws.Rows.RowHeight = 25
        cntrow = ws.Cells(ws.Rows.count, 1).End(xlUp).row
                For i = irow To cntrow
                 pn = ws.Cells(i, pncol).Value
                 picpath = ipath & "\" & pn & ".jpg"
                If KCL.isExists(picpath) = False Then
                   ws.Cells(i, Piccol).Value = "缺少图片" & picpath
                Else
                    ws.Cells(i, Piccol).Value = ""
                    Set iPic = ws.Shapes.AddPicture(picpath, False, True, 0, 0, -1, -1)
                    Set rng = ws.Cells(i, Piccol)
                With iPic ' 保持纵横比
                    .LockAspectRatio = msoTrue  '计算最佳尺寸以适应单元格
                    Dim cellRatio, imgRatio
                    cellRatio = rng.Width / rng.Height
                    imgRatio = .Width / .Height
                    If imgRatio > cellRatio Then  '图片较宽，以单元格宽度为准
                        .Width = rng.Width * 0.9
                        .Top = rng.Top + (rng.Height - .Height) / 2 '垂直居中
                        .Left = rng.Left
                    Else '图片较高，以单元格高度为准
                        .Height = rng.Height * 0.8
                        .Top = rng.Top + (rng.Height - .Height) / 2
                        .Left = rng.Left + (rng.Width - .Width) / 2 '水平居中
                    End If
                    .Placement = xlMoveAndSize
                End With
            End If
    Next i
End Sub

Public Sub xlshow()
  xlAPP.Visible = True
      With xlAPP.Windows(1)
        .WindowState = xlNormal
        .Width = 1080  ' 设置窗口宽度
        .Height = .Width * 0.618 ' 设置窗口高度
    End With
       With xlAPP
        .ScreenUpdating = True
        .Calculation = xlCalculationAutomatic 'xlCalculationmanual
        .EnableEvents = True
    End With
End Sub
  
Public Sub xlhide()
    With xlAPP
        .ScreenUpdating = False
        .Calculation = xlCalculationManual 'xlCalculationmanual
        .EnableEvents = False
    End With
 xlAPP.Visible = False
End Sub
 

