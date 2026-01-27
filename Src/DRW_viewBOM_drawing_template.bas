Attribute VB_Name = "DRW_viewBOM_drawing_template"
Private Const mdlname As String = "DRW_viewBOM_drawing_template"
Sub main()
    CATIA.RefreshDisplay = False
    On Error Resume Next


   fmt = Array("Number", "Part Number", "Quantity", "Nomenclature", "Definition", "Material", "Product Description")

    Dim tmpPath: tmpPath = CATIA.SystemService.Environ("TEMP") & "\bom_temp.md"
    Debug.Print tmpPath
    
    ' 2. 获取视图与产品
    Dim oSheet: Set oSheet = CATIA.ActiveDocument.sheets.ActiveSheet
    Dim oView: Set oView = oSheet.Views.ActiveView
    Dim oprd: Set oprd = oView.GenerativeBehavior.Document
    
    If oprd Is Nothing Then MsgBox "没有关联的 Product": Exit Sub
    
    ' 3. 导出并处理数据
    oprd.GetItem("BillOfMaterial").SetSecondaryFormat fmt
    oprd.GetItem("BillOfMaterial").Print "TXT", tmpPath, oprd
    
    Dim flatData: flatData = GetSortedBOM(tmpPath) ' 获取处理好并排序的数据
    If IsEmpty(flatData) Then Exit Sub
    
    ' 4. 更新表格
    Dim tbl, R, c
    Err.Clear: Set tbl = oView.Tables.item("GenBOM")
    If Err.Number = 0 Then oView.Selection.Clear: oView.Selection.Add tbl: oView.Selection.Delete
    
    Set tbl = oView.Tables.Add(50, 50, UBound(flatData, 1), UBound(flatData, 2), 10, 20)
    tbl.Name = "GenBOM"
    
    For R = 1 To UBound(flatData, 1)
        For c = 1 To UBound(flatData, 2)
            tbl.SetCellString R, c, CStr(flatData(R, c))
        Next
    Next
    CATIA.RefreshDisplay = True
End Sub

' --- 核心处理函数 (读取 + 解析 + 排序) ---
Function GetSortedBOM(fPath)
    Dim fso, ts, rawLines, i, validRows(), vCount, line
    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FileExists(fPath) Then Exit Function
    
    ' 读取所有行
    Set ts = fso.OpenTextFile(fPath, 1): rawLines = Split(ts.ReadAll, vbCrLf): ts.Close
    
    ' 筛选有效行 (Recapitulation 之后, 以 | 开头但不是非 +-- 行)
    ReDim validRows(UBound(rawLines))
    Dim isBody: isBody = False
    vCount = 0
    
    For i = 0 To UBound(rawLines)
        line = Trim(rawLines(i))
        If InStr(line, "Recapitulation") > 0 Then isBody = True
        If isBody And Left(line, 1) = "|" And InStr(line, "+--") = 0 Then
            validRows(vCount) = SplitLine(line): vCount = vCount + 1
        End If
    Next
    
    If vCount = 0 Then Exit Function
    
    ' 转为 2D 矩阵 (1-based fit for CATIA Table)
    Dim R, c, colCount, mat
    colCount = UBound(validRows(0)) + 1
    ReDim mat(vCount, colCount)
    
    For R = 0 To vCount - 1
        For c = 0 To colCount - 1
            mat(R + 1, c + 1) = validRows(R)(c)
        Next
    Next
    
    ' 冒泡排序 (从第2行开始，跳过标题)
    Dim j, k, tmp
    For R = 2 To vCount
        For j = R + 1 To vCount
            ' 比较第1列 (Number)
            If ShouldSwap(mat(R, 1), mat(j, 1)) Then
                For k = 1 To colCount ' 交换整行
                    tmp = mat(R, k): mat(R, k) = mat(j, k): mat(j, k) = tmp
                Next
            End If
        Next
    Next
    GetSortedBOM = mat
End Function

' --- 辅助：行分割 ---
Function SplitLine(s)
    s = Mid(s, 2, Len(s) - 2) ' 去头尾 |
    Dim arr: arr = Split(s, "|")
    Dim i: For i = 0 To UBound(arr): arr(i) = Trim(arr(i)): Next
    SplitLine = arr
End Function

' --- 辅助：排序逻辑 (数字优先 > 文本 > 空值垫底) ---
Function ShouldSwap(v1, v2)
    ShouldSwap = False
    Dim e1: e1 = (v1 = "")
    Dim e2: e2 = (v2 = "")
    
    If e1 And Not e2 Then ShouldSwap = True: Exit Function ' v1空，往后排
    If e2 Then Exit Function                               ' v2空，不动（v1在v2前）
    
    If IsNumeric(v1) And IsNumeric(v2) Then
        If CDbl(v1) > CDbl(v2) Then ShouldSwap = True
    Else
        If StrComp(v1, v2, 1) = 1 Then ShouldSwap = True
    End If
End Function

' ------------------------------------------------------------------
' Sub: 二维数组排序 (Bubble Sort)
' sortCol: 排序列索引 (1-based)
' hasHeader: 是否包含标题行
' ------------------------------------------------------------------
