Attribute VB_Name = "MDL_setThreadcolor"
'自动根据螺纹直针对螺纹孔和修饰螺纹赋予特定颜色，方便工程师直观检查。
'{Gp:4}
'{Ep:SetThreadColors}
'{Caption:螺纹孔上色}
'{ControlTipText:此按钮将零件螺纹孔上色}
'{BackColor:}

Option Explicit
Private Type Threadspec
    MinDia As Double                             ' 最小直径 (包含)
    MaxDia As Double                             ' 最大直径 (包含)
    R As Integer                                 ' 红色通道 (0-255)
    G As Integer                                 ' 绿色通道 (0-255)
    B As Integer                                 ' 蓝色通道 (0-255)
End Type

Private oSpec() As Threadspec
Private m_oSelection As Selection
Private Const MODULE_NAME As String = "MDL_setThreadcolor"
Sub SetThreadColors()
    On Error GoTo ErrorHandler
    Dim oCatia As Object
    Set oCatia = CATIA
    If oCatia.Documents.count = 0 Then
        MsgBox "请先打开一个 Part 或 Product 文档。", vbExclamation, MODULE_NAME
        Exit Sub
    End If
    Dim oDoc As Document
    Set oDoc = oCatia.ActiveDocument
    Set m_oSelection = oDoc.Selection
    oCatia.DisplayFileAlerts = False
    oCatia.RefreshDisplay = False
    Call InitConfig
    Select Case TypeName(oDoc)
    Case "PartDocument"
        Call ProcessPart(oDoc.part)
    Case "ProductDocument"
        Call ProcessProduct(oDoc.Product)
    Case Else
        MsgBox "此宏仅能在 Part 或 Product 环境下运行。", vbExclamation, MODULE_NAME
    End Select
Cleanup:
    oCatia.RefreshDisplay = True
    oCatia.DisplayFileAlerts = True
    MsgBox "螺纹染色处理完成！", vbInformation, MODULE_NAME
    Exit Sub
ErrorHandler:
    MsgBox "运行期间发生意外错误: " & Err.Description, vbCritical, MODULE_NAME
    Resume Cleanup
End Sub

Private Sub InitConfig()
    ReDim oSpec(3)
   ' 规则 0: M4 (公称直径 4.0, 范围 3.6 ~ 4.4) -> 黄色 (Yellow)
   With oSpec(0)
        .MinDia = 3.6: .MaxDia = 4.4
        .R = 255: .G = 255: .B = 0
    End With
    ' 规则 1: M5 (公称直径 5.0, 范围 4.6 ~ 5.4) -> 紫色 (Purple)
    With oSpec(1)
        .MinDia = 4.6: .MaxDia = 5.4
        .R = 255: .G = 0: .B = 255
    End With
    ' 规则 2: M6 (公称直径 6.0, 范围 5.6 ~ 6.4) -> 绿色 (Green)
    With oSpec(2)
        .MinDia = 5.6: .MaxDia = 6.4
        .R = 0: .G = 255: .B = 0
    End With
    ' 规则 3: M8及以上 (公称直径 8.0+, 范围 7.6 ~ 100) -> 蓝色 (Blue)
    With oSpec(3)
        .MinDia = 7.6: .MaxDia = 100
        .R = 0: .G = 0: .B = 255
    End With
End Sub
Private Sub ProcessProduct(ByVal oProd As Product)
    On Error GoTo ErrorHandler
    Dim i As Integer
    Dim childCount As Integer
    childCount = oProd.Products.count
    If childCount = 0 Then
        Dim oPart As part
        Set oPart = TryGetPartFromProduct(oProd)
        If Not oPart Is Nothing Then
            On Error GoTo ErrorHandler
            oProd.ApplyWorkMode 2  ' 2 = DESIGN_MODE (应用设计模式以加载零件完整特征)
            Call ProcessPart(oPart)
        End If
    Else
        For i = 1 To childCount
            Call ProcessProduct(oProd.Products.item(i))
        Next i
    End If
    Exit Sub
ErrorHandler:
    Err.Clear
End Sub

Private Sub ProcessPart(ByVal oPart As part)
    Dim oBody As body
    Dim oShape As Object
    Dim i As Integer, j As Integer
    Dim dThreadDia As Double
    Dim bIsThread As Boolean
    Dim R As Integer, G As Integer, B As Integer
    For i = 1 To oPart.bodies.count
        Set oBody = oPart.bodies.item(i)
        For j = 1 To oBody.Shapes.count
            Set oShape = oBody.Shapes.item(j)
            bIsThread = False
            dThreadDia = 0#
            Select Case TypeName(oShape)
            Case "Hole"
                If oShape.ThreadingMode = 0 Then ' 0 = catThreadedHoleThreading (已打钩开启螺纹属性)
                    If TryGetHoleThreadDiameter(oShape, dThreadDia) Then
                        bIsThread = True
                    End If
                End If
                    
            Case "Thread"
                If TryGetThreadDiameter(oShape, dThreadDia) Then
                    bIsThread = True
                End If
            End Select
            
            If bIsThread Then
                If GetColorByDia(dThreadDia, R, G, B) Then
                    ApplyColorToShape oShape, R, G, B
                End If
            End If
        Next j
    Next i
    
    m_oSelection.Clear
End Sub
Private Function TryGetPartFromProduct(ByVal oProd As Product) As part
    On Error GoTo Fail
    Set TryGetPartFromProduct = oProd.ReferenceProduct.Parent.part
    Exit Function
Fail:
    Set TryGetPartFromProduct = Nothing
    Err.Clear
End Function
Private Function TryGetHoleThreadDiameter(ByVal oHole As Object, ByRef outDiameter As Double) As Boolean
    On Error GoTo Fail
        outDiameter = oHole.ThreadDiameter.Value
        TryGetHoleThreadDiameter = True
    Exit Function
Fail:
    TryGetHoleThreadDiameter = False
    Err.Clear
End Function
Private Function TryGetThreadDiameter(ByVal oThread As Object, ByRef outDiameter As Double) As Boolean
    On Error GoTo Fail
    outDiameter = oThread.Diameter
    TryGetThreadDiameter = True
    Exit Function
Fail:
    TryGetThreadDiameter = False
    Err.Clear
End Function
Private Function GetColorByDia(ByVal dDia As Double, ByRef outR As Integer, ByRef outG As Integer, ByRef outB As Integer) As Boolean
    Dim k As Integer
    GetColorByDia = False
    For k = LBound(oSpec) To UBound(oSpec)
        If dDia >= oSpec(k).MinDia And dDia <= oSpec(k).MaxDia Then
            outR = oSpec(k).R
            outG = oSpec(k).G
            outB = oSpec(k).B
            GetColorByDia = True
            Exit Function
        End If
    Next k
End Function
' 将指定颜色应用到特征对象
Private Sub ApplyColorToShape(ByVal oShape As Object, ByVal R As Integer, ByVal G As Integer, ByVal B As Integer)
    On Error GoTo Fail
    m_oSelection.Clear
    m_oSelection.Add oShape
    m_oSelection.VisProperties.SetRealColor R, G, B, 1
Fail:
    Err.Clear
End Sub


