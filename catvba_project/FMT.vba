' ���ñ���ʽ��������
Sub FMT(Optional ByVal startCell As String = "A1", _
        Optional ByVal endCell As String = "O30", _
        Optional ByVal headerOnly As Boolean = False)
    
    Dim ws As Worksheet
    Set ws = ThisWorkbook.ActiveSheet
    
    ' ��ȡ���Χ
    Dim tableRange As String
    tableRange = startCell & ":" & endCell
    
    ' ��ȡ��ͷ��Χ
    Dim headerRange As String
    headerRange = Left(startCell, 1) & "1:" & Left(endCell, 1) & "1"
    
    ' ���ñ�ͷ��ʽ
    FormatHeader ws, headerRange
    
    ' ������ǽ���ʽ����ͷ��������������ʽ
    If Not headerOnly Then
        ' ���ý����б���ɫ
        FormatAlternateColumns ws, startCell, endCell
        
        ' ���ñ��߿�
        FormatTableBorders ws, tableRange
        
        ' ���ô�����ͼ
        SetWindowView
    End If
End Sub

' ��ʽ����ͷ
Private Sub FormatHeader(ws As Worksheet, headerRange As String)
    With ws.Range(headerRange)
        ' ������б߿�������±߿�
        .Borders.LineStyle = xlNone
        With .Borders
            .LineStyle = xlDouble
            .Weight = xlThick
            .ColorIndex = xlAutomatic
        End With
        ' �����ڲ�����
        .Borders(xlInsideVertical).LineStyle = xlContinuous
    End With
End Sub

' ���ý����б���ɫ
Private Sub FormatAlternateColumns(ws As Worksheet, startCell As String, endCell As String)
    Dim startCol As Integer, endCol As Integer
    startCol = Range(startCell).Column
    endCol = Range(endCell).Column
    
    ' ������Ҫ��ʽ�����к�����
    Dim colNumbers As New Collection
    Dim i As Integer
    For i = startCol To endCol
        If i Mod 2 = 0 Then ' ����ż���б���ɫ
            colNumbers.Add i
        End If
    Next i
    
    ' Ӧ�ø�ʽ
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

' ���ñ��߿�
Private Sub FormatTableBorders(ws As Worksheet, tableRange As String)
    With ws.Range(tableRange)
        .Borders.LineStyle = xlContinuous
        .Borders.Weight = xlThin
        .Borders.ColorIndex = 0
    End With
End Sub

' ���ô�����ͼ
Private Sub SetWindowView()
    With ActiveWindow
        .Zoom = 85
        .ScrollColumn = 8
        .ScrollRow = 10
    End With
End Sub