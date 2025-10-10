' 主函数：计算零件长宽高
Sub 零件长宽高()
    ' 检查文档类型
    If Not ValidateActiveDocument() Then Exit Sub
    
    ' 关闭同步以提高性能
    CATIA.HSOSynchronized = False
    
    ' 获取活动文档和部件
    Dim Doc As Document
    Set Doc = CATIA.ActiveDocument
    Dim part As Part
    Set part = Doc.Part
    
    ' 获取主几何体
    Dim MBD As Body
    Set MBD = part.Bodies.Item(1)
    
    ' 创建边界框混合体
    Dim BBX As HybridBody
    Set BBX = CreateBoundingBox(part)
    
    ' 提取几何体
    Dim Extract As HybridShapeExtract
    Set Extract = CreateExtract(part, MBD, BBX)
    
    ' 计算极值点
    Dim extremums As Collection
    Set extremums = CalculateExtremums(part, Extract, BBX)
    
    ' 测量距离并创建参数
    Dim dimensions As Variant
    dimensions = MeasureAndCreateParameters(part, extremums, BBX)
    
    ' 创建边界框几何体
    CreateBoundingBoxGeometry part, BBX, dimensions, extremums
    
    ' 设置显示属性
    SetDisplayProperties Doc, BBX
    
    ' 恢复同步
    CATIA.HSOSynchronized = True
    
    MsgBox "零件长宽高属性已添加"
End Sub

' 函数1：验证活动文档类型
Function ValidateActiveDocument() As Boolean
    ValidateActiveDocument = True
    If TypeName(CATIA.ActiveDocument) <> "PartDocument" Then
        MsgBox ("请确认当前活动文档是需要创建边界框的文档")
        ValidateActiveDocument = False
    End If
End Function

' 函数2：创建边界框混合体
Function CreateBoundingBox(part As Part) As HybridBody
    Dim BBX As HybridBody
    Set BBX = part.HybridBodies.Add
    BBX.Name = "Boundingbox"
    Set CreateBoundingBox = BBX
End Function

' 函数3：提取几何体
Function CreateExtract(part As Part, MBD As Body, BBX As HybridBody) As HybridShapeExtract
    Dim HSF As HybridShapeFactory
    Set HSF = part.HybridShapeFactory
    
    Dim ref As Reference
    Set ref = part.CreateReferenceFromObject(MBD)
    
    Dim Extract As HybridShapeExtract
    Set Extract = HSF.AddNewExtract(ref)
    BBX.AppendHybridShape Extract
    Extract.Compute
    
    ' 添加到选择集以便后续操作
    CATIA.ActiveDocument.Selection.Add Extract
    
    Set CreateExtract = Extract
End Function

' 函数4：计算极值点
Function CalculateExtremums(part As Part, Extract As HybridShapeExtract, BBX As HybridBody) As Collection
    Dim HSF As HybridShapeFactory
    Set HSF = part.HybridShapeFactory
    
    Dim ref As Reference
    Set ref = part.CreateReferenceFromObject(Extract)
    
    ' 创建方向向量
    Dim xDir As HybridShapeDirection
    Dim yDir As HybridShapeDirection
    Dim zDir As HybridShapeDirection
    Set xDir = HSF.AddNewDirectionByCoord(1, 0, 0)
    Set yDir = HSF.AddNewDirectionByCoord(0, 1, 0)
    Set zDir = HSF.AddNewDirectionByCoord(0, 0, 1)
    
    ' 计算各方向的极值点
    Dim xmax As HybridShapeExtremum
    Dim xmin As HybridShapeExtremum
    Dim ymax As HybridShapeExtremum
    Dim ymin As HybridShapeExtremum
    Dim zmax As HybridShapeExtremum
    Dim zmin As HybridShapeExtremum
    
    Set xmax = HSF.AddNewExtremum(ref, xDir, 1)
    BBX.AppendHybridShape xmax
    CATIA.ActiveDocument.Selection.Add xmax
    
    Set xmin = HSF.AddNewExtremum(ref, xDir, 0)
    BBX.AppendHybridShape xmin
    CATIA.ActiveDocument.Selection.Add xmin
    
    Set ymax = HSF.AddNewExtremum(ref, yDir, 1)
    BBX.AppendHybridShape ymax
    CATIA.ActiveDocument.Selection.Add ymax
    
    Set ymin = HSF.AddNewExtremum(ref, yDir, 0)
    BBX.AppendHybridShape ymin
    CATIA.ActiveDocument.Selection.Add ymin
    
    Set zmax = HSF.AddNewExtremum(ref, zDir, 1)
    BBX.AppendHybridShape zmax
    CATIA.ActiveDocument.Selection.Add zmax
    
    Set zmin = HSF.AddNewExtremum(ref, zDir, 0)
    BBX.AppendHybridShape zmin
    CATIA.ActiveDocument.Selection.Add zmin
    
    ' 更新部件
    part.Update
    
    ' 将所有极值点存储在集合中返回
    Dim extremums As Collection
    Set extremums = New Collection
    extremums.Add xmax, "xmax"
    extremums.Add xmin, "xmin"
    extremums.Add ymax, "ymax"
    extremums.Add ymin, "ymin"
    extremums.Add zmax, "zmax"
    extremums.Add zmin, "zmin"
    
    Set CalculateExtremums = extremums
End Function

' 函数5：测量距离并创建参数
Function MeasureAndCreateParameters(part As Part, extremums As Collection, BBX As HybridBody) As Variant
    Dim Doc As Document
    Set Doc = CATIA.ActiveDocument
    
    ' 获取SPA工作台用于测量
    Dim WB As Workbench
    Set WB = Doc.GetWorkbench("SPAWorkbench")
    
    ' 定义变量
    Dim Mes(2) As Measurable
    Dim Arr(5) As Double
    Dim DisX As Double, DisY As Double, DisZ As Double
    Dim xmaxc As Double, xminc As Double
    Dim ymaxc As Double, yminc As Double
    Dim zmaxc As Double, zminc As Double
    
    ' 测量X方向距离
    Set Mes(0) = WB.GetMeasurable(extremums("xmax"))
    Mes(0).GetMinimumDistancePoints extremums("xmin"), Arr
    DisX = Abs(Arr(3) - Arr(0))
    xmaxc = Arr(0): xminc = Arr(3)
    
    ' 测量Y方向距离
    Set Mes(1) = WB.GetMeasurable(extremums("ymax"))
    Mes(1).GetMinimumDistancePoints extremums("ymin"), Arr
    DisY = Abs(Arr(4) - Arr(1))
    ymaxc = Arr(1): yminc = Arr(4)
    
    ' 测量Z方向距离
    Set Mes(2) = WB.GetMeasurable(extremums("zmax"))
    Mes(2).GetMinimumDistancePoints extremums("zmin"), Arr
    DisZ = Abs(Arr(5) - Arr(2))
    zmaxc = Arr(2): zminc = Arr(5)
    
    ' 创建参数化尺寸
    Dim product2 As Part
    Set product2 = Doc.Part
    Dim parameters1 As Parameters
    Set parameters1 = product2.Parameters 'UserRefProperties
    
    Dim length1 As Length, length2 As Length, length3 As Length
    Set length1 = parameters1.CreateDimension("X向", "LENGTH", DisX)
    Set length2 = parameters1.CreateDimension("Y向", "LENGTH", DisY)
    Set length3 = parameters1.CreateDimension("Z向", "LENGTH", DisZ)
    
    ' 返回测量结果数组
    Dim results(8) As Variant
    results(0) = DisX: results(1) = DisY: results(2) = DisZ
    results(3) = xmaxc: results(4) = xminc
    results(5) = ymaxc: results(6) = yminc
    results(7) = zmaxc: results(8) = zminc
    
    MeasureAndCreateParameters = results
End Function

' 函数6：创建边界框几何体
Sub CreateBoundingBoxGeometry(part As Part, BBX As HybridBody, dimensions As Variant, extremums As Collection)
    Dim HSF As HybridShapeFactory
    Set HSF = part.HybridShapeFactory
    
    ' 提取坐标值
    Dim DisX As Double, DisY As Double, DisZ As Double
    Dim xmaxc As Double, xminc As Double
    Dim ymaxc As Double, yminc As Double
    Dim zmaxc As Double, zminc As Double
    DisX = dimensions(0): DisY = dimensions(1): DisZ = dimensions(2)
    xmaxc = dimensions(3): xminc = dimensions(4)
    ymaxc = dimensions(5): yminc = dimensions(6)
    zmaxc = dimensions(7): zminc = dimensions(8)
    
    ' 创建方向向量
    Dim xDir As HybridShapeDirection
    Dim yDir As HybridShapeDirection
    Dim zDir As HybridShapeDirection
    Set xDir = HSF.AddNewDirectionByCoord(1, 0, 0)
    Set yDir = HSF.AddNewDirectionByCoord(0, 1, 0)
    Set zDir = HSF.AddNewDirectionByCoord(0, 0, 1)
    
    ' 创建第一个点
    Dim p1 As HybridShapePointCoord
    Set p1 = HSF.AddNewPointCoord(xmaxc, yminc, zminc)
    BBX.AppendHybridShape p1
    p1.Compute
    CATIA.ActiveDocument.Selection.Add p1
    
    ' 创建第二个点
    Dim p2 As HybridShapePointCoord
    Set p2 = HSF.AddNewPointCoord(xminc, yminc, zminc)
    BBX.AppendHybridShape p2
    p2.Compute
    CATIA.ActiveDocument.Selection.Add p2
    
    ' 创建线段
    Dim ln As HybridShapeLinePtPt
    Set ln = HSF.AddNewLinePtPt(p1, p2)
    BBX.AppendHybridShape ln
    ln.Compute
    CATIA.ActiveDocument.Selection.Add ln
    
    ' 沿Y方向拉伸
    Dim ext As HybridShapeExtrude
    Set ext = HSF.AddNewExtrude(ln, DisY, 0, yDir)
    BBX.AppendHybridShape ext
    ext.Compute
    CATIA.ActiveDocument.Selection.Add ext
    
    ' 创建表面边界
    Dim bound As HybridShapeBoundary
    Set bound = HSF.AddNewBoundaryOfSurface(ext)
    BBX.AppendHybridShape bound
    bound.Compute
    CATIA.ActiveDocument.Selection.Add bound
    
    ' 沿Z方向拉伸
    Dim ext2 As HybridShapeExtrude
    Set ext2 = HSF.AddNewExtrude(bound, DisZ, 0, zDir)
    BBX.AppendHybridShape ext2
    ext2.Compute
    CATIA.ActiveDocument.Selection.Add ext2
    
    ' 平移操作
    Dim trans As HybridShapeTransform
    Set trans = HSF.AddNewTranslate(ext, zDir, DisZ)
    BBX.AppendHybridShape trans
    trans.Compute
    CATIA.ActiveDocument.Selection.Add trans
    
    ' 合并操作
    Dim asm As HybridShapeAssemble
    Set asm = HSF.AddNewJoin(ext, ext2)
    asm.AddElement trans
    BBX.AppendHybridShape asm
    asm.Compute
    
    ' 创建数据集
    Dim eles As Variant
    eles = HSF.AddNewDatums(asm)
    BBX.AppendHybridShape eles(0)
    eles(0).Name = "Bounding box of " & part.Bodies.Item(1).Name
    HSF.DeleteObjectForDatum asm
    
    ' 清理选择集
    CATIA.ActiveDocument.Selection.Delete
    CATIA.ActiveDocument.Selection.Add eles(0)
End Sub

' 函数7：设置显示属性
Sub SetDisplayProperties(Doc As Document, BBX As HybridBody)
    ' 设置透明度
    Doc.Selection.VisProperties.SetRealOpacity 100, 1
    Doc.Selection.Clear
End Sub