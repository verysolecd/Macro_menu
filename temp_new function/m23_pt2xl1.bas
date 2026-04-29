Attribute VB_Name = "m23_pt2xl1"
'Attribute VB_Name = "m23_pt2xl"
' 点坐标的导出
'{GP:2}
'{EP:pt2xl}
'{Caption:批量点坐标}
'{ControlTipText: 提示选择几何图形集后导出下面的点集}
'{BackColor:12648447}

Sub pt2xl()

    If CATIA.Windows.Count < 1 Then
        MsgBox "没有打开的窗口"
        Exit Sub
    End If
    
    Dim oDoc
    On Error Resume Next
        Set oDoc = CATIA.ActiveDocument
    On Error GoTo 0
    Dim str
    str = TypeName(oDoc)
    If Not str = "PartDocument" Then
    MsgBox "没有打开的part"
    Exit Sub
    End If
    
    
    Dim HSF:  Set HSF = oDoc.Part.HybridShapeFactory
    Dim HBS: Set HBS = oDoc.Part.HybridBodies
    Dim osel: Set osel = oDoc.Selection
    osel.Clear
    
    '=======要求选择点集和坐标
    Dim imsg, filter(0)
    imsg = "请选择点所在的几何图形集"
    filter(0) = "HybridBody"
    
    
    Set oHb = mysel(imsg, filter).Value
    
    
    Dim oAxi
    imsg = "请再选择坐标系,若无请按ESC"
    filter(0) = "AxisSystem"
  
    Set oAxi = mysel(imsg, filter).Value
    
    
    If Not oHb Is Nothing Then
        Dim i, irow, ct
        
        Set oshapes = oHb.HybridShapes
        ct = oshapes.Count
        
        ReDim Arr(0 To ct, 0 To 4)
        irow = 0
        Arr(irow, 0) = "序号"
        Arr(irow, 1) = "名称"
        Arr(irow, 2) = "X"
        Arr(irow, 3) = "Y"
        Arr(irow, 4) = "Z"
        
        irow = 1
        
        ReDim fincoord(2), absCoord(2)
        
        For i = 1 To ct
            Set oPt = oshapes.item(i)
   
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
        MsgBox "缺少待操作几何图形集，请检查选择"
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

Function mysel(prompt, filter())
    Dim osel
    Set osel = CATIA.ActiveDocument.Selection
    osel.Clear
    Dim iType(0)
'
'    Dim status
'    status = osel.SelectElement2(filter, prompt, False)

    If osel.SelectElement2(filter, prompt, False) = "Normal" Then
        If osel.Count = 1 Then
            Set mysel = osel.item(1)
        End If
    Else
    Set mysel = Nothing
    End If
    osel.Clear
End Function
' = 0 , Unknown
' = 1 , Point
' = 2 , Curve
' = 3 , Line
' = 4 , Circle
' = 5 , Surface
' = 6 , Plane
' = 7 , Solid, Volume
