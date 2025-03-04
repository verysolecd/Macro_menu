VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Class_XLM"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'version 2025.0303
' Class module for excel operations
' Public Class XLm
Option Explicit
' 定义私有变量来存储 Excel 应用程序对象和工作表对象
Public xlApp, wb, ws
Private head_bom As Variant
Private head_rv As Variant
Private inject_cols As Variant
Private extract_cols As Variant
Private bom_cols As Variant
Private startRow As Variant
' 初始化类的方法
Public Sub init()
   On Error Resume Next
    If xlApp Is Nothing Then
        Set xlApp = CreateObject("Excel.Application")
        xlApp.Visible = True
    End If
    If wb Is Nothing Then
        Set wb = xlApp.Workbooks.Add
    End If
        Set ws = wb.ActiveSheet  ' 获取sheet1
        ws.Name = "bom"
    If Err.Number <> 0 Then
        Err.Clear
    End If
    On Error GoTo 0
    head_rv = Array("Part Number", "Description", "Quantity", "Level")
    inject_cols = Array(2, 4, 6, 12, 12, 8, 13, 10)
    extract_cols = Array(1, 3, 5, 7, 9, 11)
    bom_cols = Array(1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15)
    
    head_bom = Array( _
    "No." & Chr(10) & "编号", _
    "Layout" & Chr(10) & "层级", _
    "PN" & Chr(10) & "零件号", _
    "Nomenclature" & Chr(10) & "英文名称", _
    "Definition" & Chr(10) & "中文名称", _
    "Picture" & Chr(10) & "图像", _
    "Quantity" & Chr(10) & "数量(PCS)", _
    "Weight" & Chr(10) & "单质量", _
    "Material" & Chr(10) & "材料", _
    "Thickness" & Chr(10) & "厚度(mm)", _
    "Density" & Chr(10) & "密度" _
        )
    startRow = 2
End Sub
Public Sub clear_table()
    ws.UsedRange.ClearContents
End Sub
Public Sub inject_data(rowNumber, data, iType)
    Dim maxCol As Long
    Dim outputArr As Variant
    Dim i As Long
    Dim iCols As Variant
    Dim iHeader As Variant
    Select Case iType
        Case "bom"
            iCols = bom_cols
            iHeader = head_bom
        Case "rv"
            iCols = inject_cols
            iHeader = head_rv
        Case Else
            Exit Sub
    End Select
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
Sub set_Header(header)
    Dim lastCol
    lastCol = UBound(header) - LBound(header) + 1
    With ws
    .Range(.Cells(1, 1), .Cells(1, lastCol)).Value = header
    End With
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