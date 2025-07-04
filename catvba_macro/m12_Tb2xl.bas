Attribute VB_Name = "m12_Tb2xl"
'Attribute VB_Name = "m12_Tb2xl"
' ͼֽ��ͼ�����������
'{GP:1}
'{EP:Tb2xl}
'{Caption:2D��񵼳�}
'{ControlTipText: ��ʾѡ��2ͼֽ�б��󵼳�}
'{BackColor:12648447}




Sub Tb2xl()

  If Not CanExecute("DrawingDocument") Then
          Exit Sub
     End If

    Dim oDoc As DrawingDocument
    Set oDoc = CATIA.ActiveDocument
    

    Dim oSht As DrawingSheet
    Set oSht = oDoc.Sheets.ActiveSheet
    
    ' set drawing drwView
    Dim oView As DrawingView
    Set oView = oSht.Views.ActiveView
    

    Dim imsg
    imsg = "��ѡ��table"
    
    Dim drwTable
    Set drwTable = KCL.SelectItem(imsg, DrawingTable)
  
    If Not drwTable Is Nothing Then
        Dim rowsNo As Long
        rowsNo = drwTable.NumberOfRows
    
        Dim colsNo As Long
        colsNo = drwTable.NumberOfColumns
        
        
        Dim i As Long, j As Long
        ReDim Arr(rowsNo - 1, colsNo - 1) As Variant
      
        For i = 1 To rowsNo
            For j = 1 To colsNo
                ' write cell content to an array item
                Arr(i - 1, j - 1) = drwTable.GetCellString(i, j)
            Next
        Next
        
        ArrayToxl Arr
    Else
    
    MsgBox "�޿ɲ����������"
    Exit Sub
    End If


End Sub


Sub ArrayToxl(arr2D() As Variant)
    Dim xlAPP As Object
    Set xlAPP = CreateObject("Excel.Application")
    Dim wbook As Object
    Set wbook = xlAPP.Workbooks.Add
    Dim rng As Object
    Set rng = wbook.Sheets(1).Range("B2")
    
    With rng.Resize(UBound(arr2D, 1) + 1, UBound(arr2D, 2) + 1)
        .Value = arr2D
        .Borders.LineStyle = xlContinuous
        .Borders.Weight = xlThin
        .Borders.ColorIndex = 0
    End With
    
    xlAPP.Visible = True
End Sub

