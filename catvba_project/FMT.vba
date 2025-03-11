' 设置表格格式的主函数
Sub FMT(Optional ByVal startCell As String = "A1", _
        Optional ByVal endCell As String = "O30", _
        Optional ByVal headerOnly As Boolean = False)
    
    Dim ws As Worksheet
    Set ws = ThisWorkbook.ActiveSheet
    Dim tableRange As String
    tableRange = startCell & ":" & endCell
    Dim headerRange As String
    headerRange = Left(startCell, 1) & "1:" & Left(endCell, 1) & "1"
    FormatHeader ws, headerRange
    If Not headerOnly Then
        FormatAlternateColumns ws, startCell, endCell
        FormatTableBorders ws, tableRange
        SetWindowView
    End If
End Sub
Private Sub FormatHeader(ws As Worksheet, headerRange As String)
    With ws.Range(headerRange)
        .Borders.LineStyle = xlNone
        With .Borders
            .LineStyle = xlDouble
            .Weight = xlThick
            .ColorIndex = xlAutomatic
        End With
        .Borders(xlInsideVertical).LineStyle = xlContinuous
    End With
End Sub

Private Sub FormatAlternateColumns(ws As Worksheet, startCell As String, endCell As String)
    Dim startCol As Integer, endCol As Integer
    startCol = Range(startCell).Column
    endCol = Range(endCell).Column
    Dim colNumbers As New Collection
    Dim i As Integer
    For i = startCol To endCol
        If i Mod 2 = 0 Then ' 设置偶数列背景色
            colNumbers.Add i
        End If
    Next i
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

Private Sub FormatTableBorders(ws As Worksheet, tableRange As String)
    With ws.Range(tableRange)
        .Borders.LineStyle = xlContinuous
        .Borders.Weight = xlThin
        .Borders.ColorIndex = 0
    End With
End Sub

Private Sub SetWindowView()
    With ActiveWindow
        .Zoom = 85
        .ScrollColumn = 8
        .ScrollRow = 10
    End With
End Sub