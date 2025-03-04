Sub FMT()
    With ThisWorkbook.ActiveSheet
        ' 设置标题行格式
        With .Range("A1:O1")
            ' 清除所有边框
            .Borders.LineStyle = xlNone
            ' 设置外边框
            With .Borders
                .LineStyle = xlDouble
                .Weight = xlThick
                .ColorIndex = xlAutomatic
            End With
            ' 内部竖线保留细线
            .Borders(xlInsideVertical).LineStyle = xlContinuous
        End With

        ' 设置奇数列背景色（参数改为列号数组）
        Dim colNumbers As Variant
        colNumbers = Array(2, 4, 6, 8, 10, 12, 14, 15) ' B,D,F,H,J,L,N,O列号
        
        Dim colsToFormat As Range
        Set colsToFormat = .Columns(colNumbers(0))
        For i = 1 To UBound(colNumbers)
            Set colsToFormat = Union(colsToFormat, .Columns(colNumbers(i)))
        Next
        With colsToFormat
            .Interior.ThemeColor = xlThemeColorDark1
            .Interior.TintAndShade = -0.2499
        End With

        ' 设置全表基础边框
        With .Range("A1:O30")
            .Borders.LineStyle = xlContinuous
            .Borders.Weight = xlThin
            .Borders.ColorIndex = 0
        End With
    End With

    ' 窗口设置
    With ActiveWindow
        .Zoom = 85
        .ScrollColumn = 8
        .ScrollRow = 10
    End With
End Sub