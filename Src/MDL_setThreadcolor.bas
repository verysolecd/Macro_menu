Attribute VB_Name = "MDL_setThreadcolor"
'{GP:4}
'{EP:Run_SetThreadColors}
'{Caption:批量点坐标}
'{ControlTipText: 提示选择几何图形集后导出下面的点集}
'{BackColor:}

Option Explicit
' 定义一个简单的结构体来存颜色规则
Private Type ThreadRule
    MinDia As Double    ' 最小直径
    MaxDia As Double    ' 最大直径
    R As Integer
    G As Integer
    B As Integer
End Type
Private Rules() As ThreadRule
Private Const mdlname As String = "MDL_setThreadcolor"
Sub SetThreadColors()
    InitConfig
    Dim oDoc As Document
    Set oDoc = CATIA.ActiveDocument
    CATIA.DisplayFileAlerts = False
    If TypeName(oDoc) = "PartDocument" Then
        ProcessPart oDoc.part
    ElseIf TypeName(oDoc) = "ProductDocument" Then
        ProcessProduct oDoc.Product
    Else
        MsgBox "请在 Part 或 Product 环境下运行。"
    End If
    
    CATIA.DisplayFileAlerts = True
    MsgBox "螺纹染色完成！", vbInformation
End Sub
Private Sub InitConfig()
    ReDim Rules(2)
    ' 规则 0: M3 - M5 (直径 2.9 ~ 5.5) -> 红色
    Rules(0).MinDia = 2.9
    Rules(0).MaxDia = 5.5
    Rules(0).R = 255: Rules(0).G = 0: Rules(0).B = 0
    ' 规则 1: M6 - M10 (直径 5.9 ~ 10.5) -> 绿色
    Rules(1).MinDia = 5.9
    Rules(1).MaxDia = 10.5
    Rules(1).R = 0: Rules(1).G = 255: Rules(1).B = 0
    ' 规则 2: M12及以上 (直径 11.9 ~ 100) -> 蓝色
    Rules(2).MinDia = 11.9
    Rules(2).MaxDia = 100
    Rules(2).R = 0: Rules(2).G = 0: Rules(2).B = 255
End Sub

Private Sub ProcessProduct(oProd As Product)
    Dim i As Integer
    If oProd.Products.count = 0 Then
        On Error Resume Next
        Dim oPart As part
        Set oPart = oProd.ReferenceProduct.Parent.part
        If Err.Number = 0 Then
            oProd.ApplyWorkMode 2 ' DESIGN_MODE
            ProcessPart oPart
        End If
        On Error GoTo 0
    Else
        For i = 1 To oProd.Products.count
            ProcessProduct oProd.Products.item(i)
        Next
    End If
End Sub

Private Sub ProcessPart(oPart As part)
    Dim oSel As Selection
    Set oSel = CATIA.ActiveDocument.Selection
    Dim oBody As body
    Dim oShape
    Dim i As Integer, j As Integer
    Dim dDia As Double
    Dim bIsThread As Boolean
    Dim R As Integer, G As Integer, B As Integer
    For i = 1 To oPart.bodies.count
        Set oBody = oPart.bodies.item(i)
        ' 遍历Body内的所有形状 (Shapes)
        For j = 1 To oBody.Shapes.count
            Set oShape = oBody.Shapes.item(j)
            bIsThread = False
            
            ' case A: 螺纹孔 (Hole)
            If TypeName(oShape) = "Hole" Then
                ' 检查是否开启了螺纹属性 (0 = catThreadedHoleThreading)
                If oShape.ThreadingMode = 0 Then
                    On Error Resume Next
                    dDia = oShape.ThreadDiameter.Value
                    If Err.Number = 0 Then bIsThread = True
                    On Error GoTo 0
                Else
                    ' 如果需要给普通光孔也上色，可以在这里处理
                    ' 比如: dDia = oShape.Diameter.Value
                End If
            ' case B: 修饰螺纹 (Thread)
            ElseIf TypeName(oShape) = "Thread" Then
                On Error Resume Next
                dDia = oShape.Diameter ' Thread对象的Diameter属性通常就是螺纹大径
                If Err.Number = 0 Then bIsThread = True
                On Error GoTo 0
            End If
            
            If bIsThread Then
                If GetColorByDia(dDia, R, G, B) Then
                    oSel.Clear
                    oSel.Add oShape
                    oSel.VisProperties.SetRealColor R, G, B, 1
                End If
            End If
        Next j
    Next i
    oSel.Clear
End Sub

Private Function GetColorByDia(D As Double, ByRef R As Integer, ByRef G As Integer, ByRef B As Integer) As Boolean
    Dim k As Integer
    GetColorByDia = False
    For k = LBound(Rules) To UBound(Rules)
        If D >= Rules(k).MinDia And D <= Rules(k).MaxDia Then
            R = Rules(k).R
            G = Rules(k).G
            B = Rules(k).B
            GetColorByDia = True
            Exit Function
        End If
    Next k
End Function
