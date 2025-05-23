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
Private inject_cols
Private extract_cols
Private bom_cols
Private wsEVT  As Ws_Evt
Private wb As Object
Private ws As Object
Private bws As Object

' 在类初始化时检查实例
Private Sub Class_Initialize()

On Error Resume Next
    If gws Is Nothing Then
        Set xlApp = CreateObject("Excel.Application")
        Set gwb = xlApp.Workbooks.Add
        Set gws = gwb.ActiveSheet
        Set wb = gwb
        Set ws = gws
        ws.Name = "bom"
    Else
        Set xlApp = GetObject("Excel.Application")
        Set wb = gwb
        Set ws = gws
    End If
     If Error.Number <> 0 Then
        Err.Clear
'         MsgBox "excel sheet 错误"
'     '    Exit Sub
     End If
     

'     xlApp.Visible = True
             With xlApp.Windows(1)
        .Width = 500  ' 设置窗口宽度
        .Height = 900 ' 设置窗口高度
    End With
     
     xlApp.Visible = False
     xlApp.Visible = True
    On Error GoTo 0
    
    
    Set wsEVT = New Ws_Evt
    wsEVT.Init wb
    
    
  iniarr
    
    
End Sub
Private Sub iniarr()


    
    inject_cols = Array(1, 3, 5, 7, 13, 9, 14, 11)
    
    mapping = Array(0, 1, 2, 3, 4, 9, 5, 0, 6, 7, 0, 6, 8, 0, 0, 0)
    extract_cols = Array(0, 2, 4, 6, 8, 10, 12)
    
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
End Sub
Private Sub set_Header(header)
    Dim lastCol
    lastCol = UBound(header) - LBound(header) + 1
    With ws
    .Range(.Cells(1, 1), .Cells(1, lastCol)).Value = header
    .Range(.Cells(1, 1), .Cells(1, lastCol)).Interior.Color = RGB(190, 190, 190)
    End With
End Sub

Public Sub clear_table()
    ws.UsedRange.ClearContents
End Sub

Public Sub inject_data(rowNb, data)
    Dim i As Long
      If ws.Cells(1, 1).Value <> "零件号" & Chr(10) & "Partnumber" Then
        Call set_Header(header_rv)
    End If
    
    Dim mapdata, osize
    mapdata = Array(0, 1, 2, 3, 4, 5, 6, 7, 8)
'    osize = UBound(mapdata)
'    Dim outputArr() As String
'    ReDim outputArr(1 To 1, 1 To osize) '转义数组
'    For i = 1 To UBound(outputArr, 2)
'        outputArr(1, i) = data(mapdata(i))
'        rng.Areas(i).Value = outputArr(1, i)
'    Next
'
    
    '====获取区域====
    Dim mapcells
        mapcells = Array(0, 1, 3, 5, 7, 9, 11, 13, 14)
    Dim rng As Range
        Set rng = Nothing
    For i = 1 To UBound(mapcells)
            ws.Cells(rowNb, mapcells(i)).Value = data(mapdata(i))
            ws.Cells(rowNb, mapcells(i)).Interior.Color = RGB(190, 190, 190)
    Next
    
    
    xlApp.Visible = True
End Sub

Public Sub inject_bom(data)
     Set ws = wb.ActiveSheet
     ws.Name = "BOM_xx"
     
     Call set_Header(header_bom)
     Dim fn
     startrow = 2
     lastRow = counter + startrow - 1
     With ws
          ws.Range(.Cells(startrow, 1), .Cells(lastRow, 17)).Value = data
        If lastRow >= 2 Then
            With ws.Range(.Cells(2, 9), .Cells(lastRow, 9))
                .Formula = "=G2*H2"  ' G列=第7列，H列=第8列
                .NumberFormat = "0.00"  ' 设置数字格式
            End With
        End If
         End With
     counter = 0
Call LvMg
   Call freesheet
   xlApp.Visible = True
End Sub


Public Sub freesheet()

Set ws = Nothing
Set gws = Nothing
Set gwb = Nothing


End Sub
 
Public Function extract_data(indRow As Long)
    Dim iCols
    iCols = extract_cols
    
    Dim temparr As Variant
    
    temparr = ws.Rows(indRow).Resize(1, 14).Value
    

    Dim outputArr As Variant
    Dim j As Long
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
    Set ws = APP.ActiveSheet
     With ws
            .Cells.ClearOutline
            .Outline.AutomaticStyles = False
            .Outline.SummaryRow = xlAbove
            .Outline.SummaryColumn = xlRight
    End With
    
    For row_num = 2 To ws.Cells(ws.Rows.Count, 1).End(xlUp).row ' 使用 xlApp.xlUp
        Dim cell_value As Variant
        cell_value = ws.Cells(row_num, 2).Value
        If Not IsEmpty(cell_value) Then
            ws.Rows(row_num).OutlineLevel = cell_value
        End If
    Next
    ' 设置对齐方式和缩进级别
    For Each cell In ws.Range("B2:B" & ws.Cells(ws.Rows.Count, 2).End(xlUp).row) ' 使用 xlApp.xlUp
        If Not IsEmpty(cell) Then
            cell.HorizontalAlignment = xlLeft ' 使用 xlApp.xlLeft
            If IsNumeric(cell.Value) Then
                cell.IndentLevel = cell.Value ' 修改缩进级别与单元格值相等
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
            lastRow = .Cells(.Rows.Count, 2).End(xlUp).row
            
        For i = 4 To lastRow
            If LVcol.Cells(i).Value <> "" Then
                .Rows(i).OutlineLevel = LVcol.Cells(i).Value ' 设置当前行的大纲级别
                LVcol.Cells(i).IndentLevel = LVcol.Cells(i).Value - 1 ' 设置当前单元格的缩进级别
            End If
        Next
    End With
    
    lastRow = ws.Cells(Rows.Count, 3).End(xlUp).row
    ws.Rows(lastRow + 3 & ":" & Rows.Count).EntireRow.Hidden = True
    Columns("Q:Q").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Selection.EntireColumn.Hidden = True
End Sub

