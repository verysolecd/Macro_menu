Attribute VB_Name = "MDL_pt2xl_abscoord"
'Attribute VB_Name = "m23_pt2xl"
' 点坐标的导出
'{GP:4}
'{EP:Mpt2xl}
'{Caption:批量点坐标}
'{ControlTipText: 提示选择几何图形集后导出下面的点集}
'{BackColor:}
'----------弹窗信息=----------------------------------
' %UI Label lbL_jpzcs  键盘造车手出品
' %UI Button btnOK  直接导出
' %UI Button btnWcoord 带相对坐标导出
' %UI Button btncancel  取消

Private mDoc, HSF, mHBS, msel
Private needtrans As Boolean

Private Const mdlname As String = "MDL_pt2xl_abscoord"
Sub Mpt2xl()
 If Not CanExecute("PartDocument") Then
        Exit Sub
    End If
Set mDoc = CATIA.ActiveDocument
Set HSF = mDoc.part.HybridShapeFactory
Set mHBS = mDoc.part.HybridBodies
Set msel = mDoc.Selection
needtrans = False
Dim oFrm: Set oFrm = KCL.newFrm("MDL_pt2xl_abscoord"): oFrm.Show
    Select Case oFrm.BtnClicked
        Case "btnOK":
            Call pt2xl(getHB())
         Case "btnWcoord":
                needtrans = True
            Call pt2xl(getHB())
         Case Else: Exit Sub
    End Select
End Sub
Function getHB()
    Dim imsg
       imsg = "请选择点所在的几何图形集"
       Dim oHb
       Set oHb = KCL.SelectItem(imsg, "HybridBody")
        Set getHB = oHb
End Function

Sub pt2xl(oHb)
    If Not oHb Is Nothing Then
        Dim i, irow, ct
        Set oshapes = oHb.HybridShapes
        ct = oshapes.count
        ReDim arr(0 To ct, 0 To 4)
        irow = 0  '获得表头
            arr(irow, 0) = "序号"
            arr(irow, 1) = "名称"
            arr(irow, 2) = "X"
            arr(irow, 3) = "Y"
            arr(irow, 4) = "Z"
        irow = 1
        ReDim fincoord(2)
        For i = 1 To ct
            Set opt = oshapes.item(i)
            Dim str
            str = HSF.GetGeometricalFeatureType(opt)
            If str = 1 Then
               Dim fakept:  Set fakept = HSF.AddNewPointCoordWithReference(0, 0, 0, opt)
                                oHb.AppendHybridShape fakept
                                mDoc.part.Update
               fakept.GetCoordinates fincoord
               If needtrans Then
                    Dim oAxi: Set oAxi = KCL.SelectItem("请选择坐标系", AxisSystem)
                    If Not oAxi Is Nothing Then fincoord = TransAxi(abscoord, oAxi)
               End If
                  msel.Clear
                  msel.Add fakept
                  msel.Delete
                  mDoc.part.Update
                arr(irow, 0) = irow
                arr(irow, 1) = opt.Name
                arr(irow, 2) = fincoord(0)
                arr(irow, 3) = fincoord(1)
                arr(irow, 4) = fincoord(2)
                irow = irow + 1
            End If
        Next
        ArrayToxl arr
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
    Set rng = wbook.sheets(1).Range("B2")
    With rng.Resize(UBound(arr2D, 1) + 1, UBound(arr2D, 2) + 1)
        .value = arr2D
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
    Dim Result(2)
    Result(0) = v(0) * xDir(0) + v(1) * xDir(1) + v(2) * xDir(2)
    Result(1) = v(0) * yDir(0) + v(1) * yDir(1) + v(2) * yDir(2)
    Result(2) = v(0) * zDir(0) + v(1) * zDir(1) + v(2) * zDir(2)
    TransAxi = Result
End Function


