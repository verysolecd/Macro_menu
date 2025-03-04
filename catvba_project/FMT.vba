Sub FMT()
    With ThisWorkbook.ActiveSheet
        ' ���ñ����и�ʽ
        With .Range("A1:O1")
            ' ������б߿�
            .Borders.LineStyle = xlNone
            ' ������߿�
            With .Borders
                .LineStyle = xlDouble
                .Weight = xlThick
                .ColorIndex = xlAutomatic
            End With
            ' �ڲ����߱���ϸ��
            .Borders(xlInsideVertical).LineStyle = xlContinuous
        End With

        ' ���������б���ɫ��������Ϊ�к����飩
        Dim colNumbers As Variant
        colNumbers = Array(2, 4, 6, 8, 10, 12, 14, 15) ' B,D,F,H,J,L,N,O�к�
        
        Dim colsToFormat As Range
        Set colsToFormat = .Columns(colNumbers(0))
        For i = 1 To UBound(colNumbers)
            Set colsToFormat = Union(colsToFormat, .Columns(colNumbers(i)))
        Next
        With colsToFormat
            .Interior.ThemeColor = xlThemeColorDark1
            .Interior.TintAndShade = -0.2499
        End With

        ' ����ȫ������߿�
        With .Range("A1:O30")
            .Borders.LineStyle = xlContinuous
            .Borders.Weight = xlThin
            .Borders.ColorIndex = 0
        End With
    End With

    ' ��������
    With ActiveWindow
        .Zoom = 85
        .ScrollColumn = 8
        .ScrollRow = 10
    End With
End Sub