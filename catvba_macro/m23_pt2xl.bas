Attribute VB_Name = "m23_pt2xl"
'Attribute VB_Name = "m23_pt2xl"
' ������ĵ���
'{GP:2}
'{EP:pt2xl}
'{Caption:����������}
'{ControlTipText: ��ʾѡ�񼸺�ͼ�μ��󵼳�����ĵ㼯}
'{BackColor:12648447}

Sub pt2xl()
    If Not CanExecute("PartDocument") Then
        Exit Sub
    End If
    Dim oDoc: Set oDoc = CATIA.ActiveDocument
    Dim HSF:  Set HSF = oDoc.Part.HybridShapeFactory
    Dim HBS: Set HBS = oDoc.Part.HybridBodies
    Dim osel: Set osel = oDoc.Selection
    
    '=======Ҫ��ѡ��㼯������
    Dim imsg
    imsg = "��ѡ������ڵļ���ͼ�μ�"
    Dim oHb
    Set oHb = KCL.SelectItem(imsg, HybridBody)
    Dim oAxi
    imsg = "����ѡ������ϵ"
    Set oAxi = KCL.SelectItem(imsg, AxisSystem)
    
    If Not oHb Is Nothing Then
        Dim i, irow, ct
        
        Set oshapes = oHb.HybridShapes
        ct = oshapes.Count
        
        ReDim Arr(0 To ct, 0 To 4)
        irow = 0
        Arr(irow, 0) = "���"
        Arr(irow, 1) = "����"
        Arr(irow, 2) = "X"
        Arr(irow, 3) = "Y"
        Arr(irow, 4) = "Z"
        
        irow = 1
        
        ReDim fincoord(2), absCoord(2)
        
        For i = 1 To ct
            Set oPt = oshapes.item(i)
            Dim str
            str = HSF.GetGeometricalFeatureType(oPt)
            If str = 1 Then
                Dim fakept
                Set fakept = HSF.AddNewPointCoordWithReference(0, 0, 0, oPt)
                oHb.AppendHybridShape fakept
                oDoc.Part.Update
               fakept.GetCoordinates absCoord
               
                  osel.Clear
                  osel.Add fakept
                  osel.Delete
                  oDoc.Part.Update
                If Not oAxi Is Nothing Then
                    fincoord = TransAxi(absCoord, oAxi)
                Else
                 fincoord = absCoord
                End If
                Arr(irow, 0) = irow
                Arr(irow, 1) = oPt.Name
                Arr(irow, 2) = fincoord(0)
                Arr(irow, 3) = fincoord(1)
                Arr(irow, 4) = fincoord(2)
                irow = irow + 1
            End If
        Next
        ArrayToxl Arr
    Else
        MsgBox "ȱ�ٴ���������ͼ�μ�������ѡ��"
        Exit Sub
    End If
End Sub

Sub ArrayToxl(arr2D() As Variant)
    Dim xlAPP
    Set xlAPP = CreateObject("Excel.Application")
    Dim wbook
    Set wbook = xlAPP.Workbooks.Add
    Dim rng
    Set rng = wbook.Sheets(1).Range("B2")
    With rng.Resize(UBound(arr2D, 1) + 1, UBound(arr2D, 2) + 1)
        .Value = arr2D
        .Borders.LineStyle = xlContinuous
        .Borders.Weight = xlThin
        .Borders.ColorIndex = 0
    End With
    xlAPP.Visible = True
End Sub
Function TransAxi(acoor As Variant, axi1) As Variant
    Dim origin(2), xDir(2), yDir(2), zDir(2)
    Dim i
    axi1.GetOrigin origin
    axi1.GetXAxis xDir
    axi1.GetYAxis yDir
    axi1.GetZAxis zDir
    Dim v(2) As Double
    For i = 0 To 2
        v(i) = acoor(i) - origin(i)
    Next
    Dim result(2)
    result(0) = v(0) * xDir(0) + v(1) * xDir(1) + v(2) * xDir(2)
    result(1) = v(0) * yDir(0) + v(1) * yDir(1) + v(2) * yDir(2)
    result(2) = v(0) * zDir(0) + v(1) * zDir(1) + v(2) * zDir(2)
    TransAxi = result
End Function
' = 0 , Unknown
' = 1 , Point
' = 2 , Curve
' = 3 , Line
' = 4 , Circle
' = 5 , Surface
' = 6 , Plane
' = 7 , Solid, Volume
