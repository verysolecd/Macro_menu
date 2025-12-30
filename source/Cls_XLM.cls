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
Private inject_cols
Private extract_cols
Private bom_cols

Private wsEVT
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
        ws.Name = "bom"
    Else
        Set xlAPP = GetObject("Excel.Application")
        Set WB = gwb
        Set ws = gws
    End If
    
If Error.Number <> 0 Then
        Err.Clear
     End If
On Error GoTo 0
    Set wsEVT = New Cls_WsEvt
    wsEVT.init WB
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
    
     Dim mapcells: mapcells = Array(0, 1, 3, 5, 7, 9, 11, 13, 14)
     
     Dim lastrow: lastrow = ws.Cells(ws.Rows.count, 2).End(xlUp).Row
     
    For rowNb = 2 To lastrow
        For i = 1 To UBound(mapcells)
                ws.Cells(rowNb, mapcells(i)).Interior.color = RGB(190, 190, 190)
        Next i
        Next rowNb
End Sub

Public Sub clear_table()
    ws.UsedRange.ClearContents
End Sub

Sub inject_ary(ByVal idata As Variant, Optional ByVal idcols As Variant, Optional ByVal idx As Variant)
 
    Call xlhide
     startrow = 2 ' 起始行
        If IsArray(idcols) And IsArray(idx) Then
            For j = 1 To UBound(idcols)
                ws.Cells(startrow, idcols(j)).Resize(UBound(idata, 1)).value = Application.index(idata, , idx(j))
            Next j
'        Else
'            startrow = 2:  lastRow = UBound(idata, 1) + startrow - 1
'            cols = UBound(idata, 2)
'              With ws
'                .Range(.Cells(startrow, 1), .Cells(lastRow, cols)).value = idata
'              End With
      End If
CleanUp:     ' 恢复Excel设置
    On Error GoTo 0
End Sub
Sub inject_bom(ByVal idata As Variant, Optional ByVal idcols As Variant, Optional ByVal idx As Variant)
     Call inject_ary(idata, idcols, idx)
         Dim bomrowcnt: bomrowcnt = UBound(idata, 1)
         startrow = 2
         lastrow = bomrowcnt + startrow - 1
          If lastrow >= 2 Then
          With ws
            With .Range(.Cells(startrow, 9), .Cells(lastrow, 9))
                .Formula = "=G2*H2"  ' G列=第7列，H列=第8列
                .NumberFormat = "0.00"  ' 设置数字格式
            End With
          End With
        End If
        Call setxlHead("bom")
        Call LvMg
End Sub

Public Sub setxlHead(ByVal s$)
Select Case s
    Case "bom": Call set_Header(header_bom, 1)
    Case "rv": Call set_Header(header_rv, 1)
    Case "gxbom": Call set_Header(hdgx, 4)
 End Select
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
Public Function extract_ary()
 Dim resultAry As Variant
 Dim cntx, cnty
    cntx = ws.Cells(ws.Rows.count, 1).End(xlUp).Row
    cnty = 14
 ReDim resultAry(1 To cntx, 1 To cnty)
 
  For idrow = 2 To cntx
         Dim temparr As Variant
            temparr = ws.Rows(idrow).Resize(1, cnty).value
        For j = 1 To cnty
            resultAry(idrow, j) = temparr(1, j)
        Next j
   Next idrow
    extract_ary = resultAry ' 返回提取的数据
End Function

Private Sub LvMg()
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
    For row_num = 2 To ws.Cells(ws.Rows.count, 2).End(xlUp).Row ' 使用 xlApp.xlUp
        Dim cell_value As Variant
        cell_value = ws.Cells(row_num, 2).value
        If Not IsEmpty(cell_value) Then
            ws.Rows(row_num).OutlineLevel = cell_value
        End If
    Next
    ' 设置对齐方式和缩进级别
    For Each cell In ws.Range("B2:B" & ws.Cells(ws.Rows.count, 2).End(xlUp).Row) ' 使用 xlApp.xlUp
        If Not IsEmpty(cell) Then
            cell.HorizontalAlignment = xlLeft ' 使用 xlApp.xlLeft
            If IsNumeric(cell.value) Then
                cell.IndentLevel = cell.value ' 修改缩进级别与单元格值相等
            End If
        End If
    Next
'   lastRow = ws.Cells(Rows.count, 3).End(xlUp).Row
'   ws.Rows(lastRow + 3 & ":" & Rows.count).EntireRow.Hidden = True
'   Columns("Q:Q").Select
'   Range(Selection, Selection.End(xlToRight)).Select
'   Selection.EntireColumn.Hidden = True
End Sub
Public Sub inject_gxbom(Data)
xlhide
     ws.Name = "BOM_xx" & Format(Now, "yymmdd_hhmm")
      startrow = 5
     Call set_Header(hdgx, startrow - 1)
     lastrow = counter + startrow - 1
     Dim cols
     cols = UBound(Data, 2)
     With ws
          ws.Range(.Cells(startrow, 1), .Cells(lastrow, cols)).value = Data
        If lastrow >= 5 Then
            With ws.Range(.Cells(startrow, 14), .Cells(lastrow, 14))
                .Formula = "=k5*m5"  ' G列=第7列，H列=第8列
                .NumberFormat = "0.00"  ' 设置数字格式
            End With
        End If
         End With
     counter = 0
xlshow
End Sub
Sub inject_pic(ipath, pncol, Piccol)
        Dim i&, rng As Range
        Dim iPic As Shape
        Dim cntrow
        ws.Rows.RowHeight = 25
        cntrow = ws.Cells(ws.Rows.count, 1).End(xlUp).Row
                For i = 2 To cntrow
                 pn = ws.Cells(i, pncol).value
                 picpath = ipath & "\" & pn & ".jpg"
                If KCL.isExists(picpath) = False Then
                   ws.Cells(i, Piccol).value = "缺少图片" & picpath
                Else
                    ws.Cells(i, Piccol).value = ""
                    '先插入图片：路径，不链接，保存图片，-1保留原有高度宽度
                    Set iPic = ws.Shapes.AddPicture(picpath, False, True, 0, 0, -1, -1)
                    Set rng = ws.Cells(i, Piccol)
                With iPic ' 保持纵横比
                    .LockAspectRatio = msoTrue  '计算最佳尺寸以适应单元格
                    Dim cellRatio, imgRatio
                    cellRatio = rng.Width / rng.Height
                    imgRatio = .Width / .Height
                    If imgRatio > cellRatio Then  '图片较宽，以单元格宽度为准
                        .Width = rng.Width * 0.9
                        .top = rng.top + (rng.Height - .Height) / 2 '垂直居中
                        .Left = rng.Left
                    Else '图片较高，以单元格高度为准
                        .Height = rng.Height * 0.8
                        .top = rng.top + (rng.Height - .Height) / 2
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
 xlAPP.Visible = fasle
End Sub
 
