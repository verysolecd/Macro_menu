VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Class_XLM"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private header_bom(1 To 17)
Private header_rv(1 To 14)
Private hdgx(1 To 15)
Private inject_cols
Private extract_cols
Private bom_cols

Private wsEVT  As Ws_Evt
Private WB As Object
Private ws As Object
Private bws As Object

' 在类初始化时检查实例
Private Sub Class_Initialize()

On Error Resume Next
    If gws Is Nothing Then
        Set xlAPP = CreateObject("Excel.Application")
        Set gwb = xlAPP.Workbooks.Add
        Set gws = gwb.ActiveSheet
        Set WB = gwb
        Set ws = gws
        ws.name = "bom"
    Else
        Set xlAPP = GetObject("Excel.Application")
        Set WB = gwb
        Set ws = gws
    End If
     If Error.Number <> 0 Then
        Err.Clear
'         MsgBox "excel sheet 错误"
'     '    Exit Sub
     End If
    On Error GoTo 0
  xlAPP.Visible = True
    With xlAPP.Windows(1)
        .WindowState = xlNormal
        .width = 500  ' 设置窗口宽度
        .height = 1080 ' 设置窗口高度
    End With
'
    xlAPP.Visible = False

    On Error GoTo 0
    Set wsEVT = New Ws_Evt
    wsEVT.Init WB
  iniarr
    
End Sub
Private Sub iniarr()
    inject_cols = Array(1, 3, 5, 7, 13, 9, 14, 11)
    mapping = Array(0, 1, 2, 3, 4, 9, 5, 0, 6, 7, 0, 6, 8, 0, 0, 0)
    'mapping = Array(1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16)
    
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
Private Sub set_Header(header, idxrow)
    Dim lastCol
    lastCol = UBound(header) - LBound(header) + 1
    With ws
    .Range(.Cells(idxrow, 1), .Cells(idxrow, lastCol)).value = header
    .Range(.Cells(idxrow, 1), .Cells(idxrow, lastCol)).Interior.color = RGB(190, 190, 190)
    End With
End Sub

Public Sub clear_table()
    ws.UsedRange.ClearContents
End Sub

Public Sub inject_data(rowNb, Data)
    ws.name = "revise prd"
    Dim i As Long
      If ws.Cells(1, 1).value <> "零件号" & Chr(10) & "Partnumber" Then
        Call set_Header(header_rv, 1)
    End If
    Dim mapdata, osize
    mapdata = Array(0, 1, 2, 3, 4, 5, 6, 7, 8)
    '====获取区域====
    Dim mapcells
        mapcells = Array(0, 1, 3, 5, 7, 9, 11, 13, 14)
    Dim rng As Range
        Set rng = Nothing
    For i = 1 To UBound(mapcells)
            ws.Cells(rowNb, mapcells(i)).value = Data(mapdata(i))
            ws.Cells(rowNb, mapcells(i)).Interior.color = RGB(190, 190, 190)
    Next
    xlAPP.Visible = True
End Sub

Public Sub inject_bom(Data)
     ws.name = "BOM_xx"
    startrow = 2
     Call set_Header(header_bom, startrow - 1)
     lastRow = counter + startrow - 1
     Dim cols
     cols = UBound(Data, 2)
     With ws
          ws.Range(.Cells(startrow, 1), .Cells(lastRow, cols)).value = Data
        If lastRow >= 2 Then
            With ws.Range(.Cells(startrow, 9), .Cells(lastRow, 9))
                .Formula = "=G2*H2"  ' G列=第7列，H列=第8列
                .NumberFormat = "0.00"  ' 设置数字格式
            End With
        End If
         End With
     counter = 0
    Call LvMg
    xlAPP.Visible = True
End Sub

Public Sub freesheet()
Set ws = Nothing
Set gws = Nothing
Set gwb = Nothing
Set xlAPP = Nothing
End Sub
 
Public Function extract_data(indRow)
    Dim iCols
    iCols = Array(0, 2, 4, 6, 8, 10, 12)
    Dim temparr As Variant
    temparr = ws.Rows(indRow).Resize(1, 14).value
    Dim j As Long
    Dim outputArr As Variant
    ReDim outputArr(1 To UBound(iCols))
    For j = 1 To UBound(iCols)
         outputArr(j) = ""
         If IsEmpty(temparr(1, iCols(j))) = False Then
         outputArr(j) = temparr(1, iCols(j))
         End If
    Next
    extract_data = outputArr ' 返回提取的数据
End Function

Private Sub LvMg()
    Dim row_num As Long ' 声明 row_num 变量
    Dim cell ' 声明 cell 变量
'    Set ws = xlApp.ActiveSheet
     With ws
            .Cells.ClearOutline
            .Outline.AutomaticStyles = False
            .Outline.SummaryRow = xlAbove
            .Outline.SummaryColumn = xlRight
    End With
    
    For row_num = 2 To ws.Cells(ws.Rows.count, 1).End(xlUp).row ' 使用 xlApp.xlUp
        Dim cell_value As Variant
        cell_value = ws.Cells(row_num, 2).value
        If Not IsEmpty(cell_value) Then
            ws.Rows(row_num).OutlineLevel = cell_value
        End If
    Next
    ' 设置对齐方式和缩进级别
    For Each cell In ws.Range("B2:B" & ws.Cells(ws.Rows.count, 2).End(xlUp).row) ' 使用 xlApp.xlUp
        If Not IsEmpty(cell) Then
            cell.HorizontalAlignment = xlLeft ' 使用 xlApp.xlLeft
            If IsNumeric(cell.value) Then
                cell.IndentLevel = cell.value ' 修改缩进级别与单元格值相等
            End If
        End If
    Next
End Sub

Public Sub LvMg2()
    Dim lastRow, i, LVcol
    With ws
            .Cells.ClearOutline
            .Outline.AutomaticStyles = False
            .Outline.SummaryRow = xlAbove
            .Outline.SummaryColumn = xlRight
            Set LVcol = .Columns(2)
            LVcol.HorizontalAlignment = xlLeft
            lastRow = .Cells(.Rows.count, 2).End(xlUp).row
            
        For i = 4 To lastRow
            If LVcol.Cells(i).value <> "" Then
                .Rows(i).OutlineLevel = LVcol.Cells(i).value ' 设置当前行的大纲级别
                LVcol.Cells(i).IndentLevel = LVcol.Cells(i).value - 1 ' 设置当前单元格的缩进级别
            End If
        Next
    End With
    lastRow = ws.Cells(Rows.count, 3).End(xlUp).row
    ws.Rows(lastRow + 3 & ":" & Rows.count).EntireRow.Hidden = True
    Columns("Q:Q").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Selection.EntireColumn.Hidden = True
End Sub

Public Sub inject_gxbom(Data)
     ws.name = "BOM_xx" & Format(Now, "yymmdd_hhmm")
      startrow = 5
     Call set_Header(hdgx, startrow - 1)
     lastRow = counter + startrow - 1
     Dim cols
     cols = UBound(Data, 2)
     With ws
          ws.Range(.Cells(startrow, 1), .Cells(lastRow, cols)).value = Data
        If lastRow >= 5 Then
            With ws.Range(.Cells(startrow, 14), .Cells(lastRow, 14))
                .Formula = "=k5*m5"  ' G列=第7列，H列=第8列
                .NumberFormat = "0.00"  ' 设置数字格式
            End With
        End If
         End With
     counter = 0
 
    xlAPP.Visible = True
End Sub

