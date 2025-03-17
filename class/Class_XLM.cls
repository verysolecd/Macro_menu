VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Class_XLM"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public wb, ws
Private header_bom(1 To 12)
Private header_rv(1 To 14)
Private inject_cols
Private extract_cols
Private bom_cols
' 在类初始化时检查实例
Private Sub Class_Initialize()
iniarr
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
     xlApp.Visible = True
     xlApp.Hide
    On Error GoTo 0
End Sub
Private Sub iniarr()
    inject_cols = Array(1, 3, 5, 7, 13, 9, 14, 11)
    extract_cols = Array(1, 3, 5, 7, 9, 11)
    bom_cols = Array(1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15)
    
    header_bom(1) = "No." & Chr(10) & "编号"
    header_bom(2) = "Layout" & Chr(10) & "层级"
    header_bom(3) = "PN" & Chr(10) & "零件号"
    header_bom(4) = "Nomenclature" & Chr(10) & "英文名称"
    header_bom(5) = "Definition" & Chr(10) & "中文名称"
    header_bom(6) = "Picture" & Chr(10) & "图像"
    header_bom(7) = "Quantity" & Chr(10) & "数量(PCS)"
    header_bom(8) = "Weight" & Chr(10) & "单质量"
    header_bom(9) = "Material" & Chr(10) & "材料"
    header_bom(10) = "Thickness" & Chr(10) & "厚度(mm)"
    header_bom(11) = "Density" & Chr(10) & "密度"

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
    End With
End Sub

Public Sub clear_table()
    ws.UsedRange.ClearContents
End Sub

Public Sub inject_data(rowNumber, data)
    Dim maxCol As Long
    Dim outputArr As Variant
    Dim i As Long
    Dim iCols As Variant
    Dim iHeader As Variant
     iCols = inject_cols
     iHeader = header_rv
    If ws.Cells(1, 1).Value = "" Then
        Call set_Header(iHeader)
    End If
    maxCol = xlApp.Max(iCols)
    ReDim outputArr(1 To 1, 1 To maxCol)  '转义数组，一行多列，其中列数由 iCols 数组的最大值决定
    ' 检查 data 数组长度是否足够
    Dim dataLength As Long
     dataLength = UBound(data) - LBound(data) + 1
    Dim maxIndex As Long
    maxIndex = xlApp.Min(dataLength, UBound(iCols) - LBound(iCols) + 1)
    For i = 1 To maxIndex
        outputArr(1, iCols(LBound(iCols) + i - 1)) = data(LBound(data) + i - 1)
    Next i
    ws.Range(ws.Cells(rowNumber, 1), ws.Cells(rowNumber, maxCol)).Value = outputArr
End Sub

Sub inject_bom(data)
     Dim fn
    startrow = 2
    lastrow = counter + startrow
    With ws
    ws.Range(.Cells(startrow, 1), .Cells(lastrow, 11)).Value = data
  End With
End Sub
    xlApp.Visible = True
    
    counter = 0




End Sub
 
Public Function extract_data(row As Long)
    Dim iCols
    iCols = extract_cols
    Dim outputArr As Variant
    Dim j As Long
    ReDim outputArr(1 To 1, 1 To UBound(iCols) - LBound(iCols) + 1)        ' 定义输出数组的大小
    For j = LBound(iCols) To UBound(iCols)     ' 从工作表中提取指定行的数据
        outputArr(1, j - LBound(iCols) + 1) = ws.Cells(row, iCols(j)).Value
    Next j
    extract_data = outputArr ' 返回提取的数据
End Function

Public Sub lvmg()
    Const xlUp = -4162
    Const xlLeft = -4131
    Dim APP
    Set APP = GetObject(, "Excel.Application")
    Dim row_num As Long ' 声明 row_num 变量
    Dim cell ' 声明 cell 变量
    ' 设置大纲级别
    For row_num = 2 To ws.Cells(ws.Rows.Count, 1).End(xlUp).row ' 使用 xlApp.xlUp
        Dim cell_value As Variant
        cell_value = ws.Cells(row_num, 2).Value
        If Not IsEmpty(cell_value) Then
            ws.Rows(row_num).OutlineLevel = cell_value - 1
        End If
    Next row_num
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


Public Sub inject_oData(odata)
     Dim rnb, cnb
     rnb = 2
     cnb = 1
    If Not IsArray(odata) Then
        MsgBox "输入参数不是数组", vbExclamation
        Exit Sub
    End If
    
    ' 获取数组的维度
    Dim dimensions As Long
    On Error Resume Next
    dimensions = UBound(odata, 2)
    On Error GoTo 0
    
    ' 确定数组的行数和列数
    Dim rowCount As Long, colCount As Long
    If dimensions > 0 Then
        ' 二维数组
        rowCount = UBound(odata, 1) - LBound(odata, 1) + 1
        colCount = UBound(odata, 2) - LBound(odata, 2) + 1
    Else
        ' 一维数组，视为单行多列
        rowCount = 1
        colCount = UBound(odata) - LBound(odata) + 1
        ' 重新定义为二维数组以便写入
        Dim tempArray
        ReDim tempArray(1 To 1, 1 To colCount)
        For i = 1 To colCount
            tempArray(1, i) = odata(LBound(odata) + i - 1)
        Next i
        odata = tempArray
    End If
    
    ' 写入数据到工作表
    ws.Range(ws.Cells(rnb, cnb), ws.Cells(rnb + rowCount - 1, cnb + colCount - 1)).Value = odata
End Sub






