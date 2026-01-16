Attribute VB_Name = "ASM_bomformat"
Option Explicit

Sub GenerateRecapBOMToTable()
    On Error GoTo ErrHandler

    '--- 1 获取 CATIA 应用 -------------------------------------------------
    Dim CATIA As Object
    Set CATIA = GetObject(, "CATIA.Application")

    '--- 2 获取根装配 ----------------------------------------------------
    Dim rootPrd As Product
    Set rootPrd = CATIA.ActiveDocument.Product

    '--- 3 获取 Bill of Material 对象 --------------------------------------
    Dim oConv As BillOfMaterial
    Set oConv = rootPrd.GetItem("BillOfMaterial")
    If oConv Is Nothing Then
        MsgBox "未能获取 BillOfMaterial 对象，请确认当前文档是装配并已生成 BOM。", vbCritical
        Exit Sub
    End If

    '--- 4 定义导出列 ----------------------------------------------------
    Dim Ary(7) As String
    Ary(0) = "Number"
    Ary(1) = "Part Number"
    Ary(2) = "Quantity"
    Ary(3) = "Nomenclature"
    Ary(4) = "Definition"
    Ary(5) = "Mass"
    Ary(6) = "Density"
    Ary(7) = "Material"
    oConv.SetSecondaryFormat Ary   ' 或 SetCurrentFormat Ary

    '--- 5 临时文件路径 ----------------------------------------------------
    Dim sTempFile As String
    sTempFile = CATIA.SystemService.Environ("TEMP") & "\bom_recap.txt"

    '--- 6 调用 Print (使用 CallByName 避免 438 错误) ------------------------
    ' 方式 A：仅两参数（最安全）
    ' 如果这行仍然报错，请确保引用了 OLE Automation 和 CATIA V5 Automation
    On Error Resume Next
    CallByName oConv, "Print", VbMethod, "TXT", sTempFile
    If Err.Number <> 0 Then
        MsgBox "导出 BOM 失败: " & Err.Description, vbCritical
        Err.Clear
        On Error GoTo 0
        Exit Sub
    End If
    On Error GoTo ErrHandler
    
    'MsgBox "BOM 已成功导出至：" & sTempFile, vbInformation
    
    '--- 7 读取并解析为二维数组 ---------------------------------------------
    Dim bomArray() As String
    bomArray = GetBOMDataArray(sTempFile)
    
    '--- 8 (测试) 打印数组内容到立即窗口 -------------------------------------
    Dim r As Integer, c As Integer
    ' 检查数组是否被初始化
    If (Not bomArray) = -1 Then
         MsgBox "BOM 内容为空或解析失败", vbExclamation
         Exit Sub
    End If

    Debug.Print "========== BOM 数组内容 =========="
    On Error Resume Next '防止越界
    For r = LBound(bomArray, 1) To UBound(bomArray, 1)
        Dim lineInfo As String: lineInfo = ""
        For c = LBound(bomArray, 2) To UBound(bomArray, 2)
            lineInfo = lineInfo & "[" & bomArray(r, c) & "] "
        Next c
        Debug.Print "Row " & r & ": " & lineInfo
    Next r
    On Error GoTo ErrHandler

    ' 可以在这里添加后续处理逻辑，比如写入 Excel 或表格

    Exit Sub

ErrHandler:
    MsgBox "错误 " & Err.Number & ": " & Err.Description, vbCritical
End Sub

'================================================================================
' 函数：读取 BOM 文本文件并解析为二维不规则数组 (仅包含数据内容)
' 返回：String() 二维数组 (行, 列)
'================================================================================
Function GetBOMDataArray(filePath As String) As String()
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    GetBOMDataArray = Split("") ' 初始化为空数组以防万一
    
    If Not fso.FileExists(filePath) Then Exit Function
    
    Dim ts As Object
    Set ts = fso.OpenTextFile(filePath, 1) ' ForReading
    
    ' 使用 Collection 暂存有效行
    Dim rawLines As New Collection
    Dim lineContent As String
    
    ' 1. 读取有效行到集合中
    Do Until ts.AtEndOfStream
        lineContent = Trim(ts.ReadLine)
        ' 过滤条件：以 "|" 开头，且不包含分隔线特征 "+-"
        If Left(lineContent, 1) = "|" Then
            ' 简单的表头判断：这里假设我们要排除标题行
            ' 如果你需要包含标题行，请注释掉下面这个 If 判断
            If InStr(lineContent, "Part Number") = 0 And InStr(lineContent, "----------") = 0 Then
               rawLines.Add lineContent
            End If
        End If
    Loop
    ts.Close
    
    If rawLines.Count = 0 Then Exit Function
    
    ' 2. 确定数组大小
    ' 先解析第一行数据来确定列数
    Dim tempArr() As String
    tempArr = ParseLineToArray(rawLines(1))
    Dim colCount As Integer
    colCount = UBound(tempArr) - LBound(tempArr) + 1
    
    Dim resultArr() As String
    ReDim resultArr(1 To rawLines.Count, 1 To colCount)
    
    ' 3. 填充二维数组
    Dim i As Integer, j As Integer
    Dim rowData() As String
    Dim maxCol As Integer
    
    For i = 1 To rawLines.Count
        rowData = ParseLineToArray(rawLines(i))
        
        ' 防止某些行格式异常，取该行实际列数与预设列数的较小值
        maxCol = UBound(rowData) - LBound(rowData) + 1
        If maxCol > colCount Then maxCol = colCount
        
        For j = 1 To maxCol
            resultArr(i, j) = rowData(j - 1) 'rowData 是 0-based
        Next j
        
        ' 填充剩余列为空字符串（如果有）
        For j = maxCol + 1 To colCount
             resultArr(i, j) = ""
        Next j
    Next i
    
    GetBOMDataArray = resultArr
End Function

' 辅助：将单行 "| A | B |" 格式解析为一维数组
Function ParseLineToArray(lineStr As String) As String()
    Dim tempStr As String
    tempStr = Trim(lineStr)
    
    ' 去首尾竖线
    If Left(tempStr, 1) = "|" Then tempStr = Mid(tempStr, 2)
    If Right(tempStr, 1) = "|" Then tempStr = Left(tempStr, Len(tempStr) - 1)
    
    ' 分割
    Dim arr() As String
    arr = Split(tempStr, "|")
    
    ' Trim 每一项
    Dim k As Integer
    For k = LBound(arr) To UBound(arr)
        arr(k) = Trim(arr(k))
    Next k
    
    ParseLineToArray = arr
End Function