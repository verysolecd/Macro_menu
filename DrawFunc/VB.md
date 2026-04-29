CATIA.RefreshDisplay = False
    On Error Resume Next
    ' 1. 定义格式与路径
    Dim fmt(6): fmt(0)="Number": fmt(1)="Part Number": fmt(2)="Quantity": fmt(3)="Nomenclature"
    fmt(4)="Definition": fmt(5)="Material"    : fmt(6)="Product Description"
    Dim tmpPath: tmpPath = CATIA.SystemService.environ("TEMP") & "\bom_temp.txt"
    ' 2. 获取视图与产品
    Dim oSheet : Set oSheet = CATIA.ActiveDocument.Sheets.ActiveSheet
    Dim oView : Set oView = oSheet.Views.ActiveView
    Dim oPrd : Set oPrd = oView.GenerativeBehavior.Document
	dim sel: set sel=CATIA.ActiveDocument.Selection
    If oPrd Is Nothing Then MsgBox "没有关联的 Product": Exit Sub
    ' 3. 导出并处理数据
    oPrd.GetItem("BillOfMaterial").SetSecondaryFormat fmt
    oPrd.GetItem("BillOfMaterial").Print "TXT", tmpPath, oPrd
    Dim flatData: flatData = GetSortedBOM(tmpPath) ' 获取处理好并排序的数据
    If IsEmpty(flatData) Then Exit Sub
    ' 4. 更新表格
	dim pox,poy :pox=0:poy=0
    Dim tbl As DrawingTable, r, c
    Err.Clear: Set tbl = oView.Tables.GetItem("GenBOM")
    If Err.Number = 0 Then
	pox=tbl.x+20
	poy=tbl.y-20
	sel.Clear: sel.Add tbl: sel.Delete	
	end if
    Set tbl = oView.Tables.Add(pox, poy, UBound(flatData, 1), UBound(flatData, 2), 10, 20)
    tbl.Name = "GenBOM"
    For r = 1 To UBound(flatData, 1)
        For c = 1 To UBound(flatData, 2)
            tbl.SetCellString r, c, CStr(flatData(r, c))
        Next
    Next
    CATIA.RefreshDisplay = True
End Sub
' --- 核心处理函数 (读取 + 解析 + 排序) ---
Function GetSortedBOM(fPath)
    Dim fso, ts, fullText, rawLines, i, validRows(), vCount, line
    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FileExists(fPath) Then Exit Function
    ' 1. 读取全文并按行分割
    Set ts = fso.OpenTextFile(fPath, 1): fullText = ts.ReadAll: ts.Close
    rawLines = Split(fullText, vbCrLf)
    ' 初始化
    ReDim validRows(UBound(rawLines))
    vCount = 0
    Dim isBody: isBody = False
    Dim targetPipes: targetPipes = 0
    Dim currentBuffer: currentBuffer = ""
    Dim bufferPipes
    For i = 0 To UBound(rawLines)
        line = Trim(rawLines(i))
        ' 定位 Recapitulation
        If InStr(line, "Recapitulation") > 0 Then isBody = True
        ' 只处理 Body 部分且忽略 +---+ 分隔线
        If isBody And InStr(line, "+--") = 0 And line <> "" Then
            ' A. 获取目标列数 (基于第一行标题行)
            If targetPipes = 0 Then
                If Left(line, 1) = "|" Then
                    ' 计算列分隔符数量 (Split 后的 UBound 即为 | 的数量，例如 |A| -> Split得3个元素, UBound=2, 即2个|)
                    targetPipes = UBound(Split(line, "|"))
                    ' 立即加入标题行
                    validRows(vCount) = SplitLine(line)
                    vCount = vCount + 1
                End If
            Else
                ' B. 数据行处理 (基于 | 数量累加)
                ' 1. 寻找新行的起始 (如果buffer空)
                If currentBuffer = "" Then
                    If Left(line, 1) = "|" Then
                        currentBuffer = line
                    End If
                Else
                    ' 2. 如果buffer不空，说明处于多行模式，追加内容
                     ' 补个换行符或空格，视需求而定。为了保留描述格式，这里用空格或换行
                    currentBuffer = currentBuffer & vbLf & line
                End If
                ' 3. 检查是否凑齐了足够的 |
                If currentBuffer <> "" Then
                    bufferPipes = UBound(Split(currentBuffer, "|"))
                    ' 如果 | 数量达标 (>= targetPipes)，说明这一行(可能跨多行)结束了
                    If bufferPipes >= targetPipes Then
                        validRows(vCount) = SplitLine(currentBuffer)
                        vCount = vCount + 1
                        currentBuffer = "" ' 清空，准备读下一行
                    End If
                End If
            End If
        End If
    Next
    If vCount = 0 Then Exit Function
    ' 转为 2D 矩阵
    Dim r, c, colCount, mat
    colCount = UBound(validRows(0)) + 1
    ReDim mat(vCount, colCount)
    For r = 0 To vCount - 1
        For c = 0 To colCount - 1
            If c <= UBound(validRows(r)) Then
                mat(r + 1, c + 1) = validRows(r)(c)
            Else
                mat(r + 1, c + 1) = ""
            End If
        Next
    Next
    ' 冒泡排序
    Dim j, k, tmp
    For r = 2 To vCount
        For j = r + 1 To vCount
            If ShouldSwap(mat(r, 1), mat(j, 1)) Then
                For k = 1 To colCount
                    tmp = mat(r, k): mat(r, k) = mat(j, k): mat(j, k) = tmp
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