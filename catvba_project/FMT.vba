' 设置表格格式的主函数
Sub FMT(Optional ByVal startCell As String = "A1", _
        Optional ByVal endCell As String = "O30", _
        Optional ByVal headerOnly As Boolean = False)
    
    Dim ws As Worksheet
    Set ws = ThisWorkbook.ActiveSheet
    
    ' 获取表格范围
    Dim tableRange As String
    tableRange = startCell & ":" & endCell
    
    ' 获取表头范围
    Dim headerRange As String
    headerRange = Left(startCell, 1) & "1:" & Left(endCell, 1) & "1"
    
    ' 设置表头格式
    FormatHeader ws, headerRange
    
    ' 如果不是仅格式化表头，则设置其他格式
    If Not headerOnly Then
        ' 设置交替列背景色
        FormatAlternateColumns ws, startCell, endCell
        
        ' 设置表格边框
        FormatTableBorders ws, tableRange
        
        ' 设置窗口视图
        SetWindowView
    End If
End Sub

' 格式化表头
Private Sub FormatHeader(ws As Worksheet, headerRange As String)
    With ws.Range(headerRange)
        ' 清除所有边框后设置新边框
        .Borders.LineStyle = xlNone
        With .Borders
            .LineStyle = xlDouble
            .Weight = xlThick
            .ColorIndex = xlAutomatic
        End With
        ' 保留内部竖线
        .Borders(xlInsideVertical).LineStyle = xlContinuous
    End With
End Sub

' 设置交替列背景色
Private Sub FormatAlternateColumns(ws As Worksheet, startCell As String, endCell As String)
    Dim startCol As Integer, endCol As Integer
    startCol = Range(startCell).Column
    endCol = Range(endCell).Column
    
    ' 创建需要格式化的列号数组
    Dim colNumbers As New Collection
    Dim i As Integer
    For i = startCol To endCol
        If i Mod 2 = 0 Then ' 设置偶数列背景色
            colNumbers.Add i
        End If
    Next i
    
    ' 应用格式
    If colNumbers.Count > 0 Then
        Dim colsToFormat As Range
        Set colsToFormat = ws.Columns(colNumbers(1))
        For i = 2 To colNumbers.Count
            Set colsToFormat = Union(colsToFormat, ws.Columns(colNumbers(i)))
        Next i
        
        With colsToFormat
            .Interior.ThemeColor = xlThemeColorDark1
            .Interior.TintAndShade = -0.2499
        End With
    End If
End Sub

' 设置表格边框
Private Sub FormatTableBorders(ws As Worksheet, tableRange As String)
    With ws.Range(tableRange)
        .Borders.LineStyle = xlContinuous
        .Borders.Weight = xlThin
        .Borders.ColorIndex = 0
    End With
End Sub

' 设置窗口视图
Private Sub SetWindowView()
    With ActiveWindow
        .Zoom = 85
        .ScrollColumn = 8
        .ScrollRow = 10
    End With
End Sub