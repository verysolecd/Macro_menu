' ��������������������
Sub ��������()
    ' ����ĵ�����
    If Not ValidateActiveDocument() Then Exit Sub
    
    ' �ر�ͬ�����������
    CATIA.HSOSynchronized = False
    
    ' ��ȡ��ĵ��Ͳ���
    Dim Doc As Document
    Set Doc = CATIA.ActiveDocument
    Dim part As Part
    Set part = Doc.Part
    
    ' ��ȡ��������
    Dim MBD As Body
    Set MBD = part.Bodies.Item(1)
    
    ' �����߽������
    Dim BBX As HybridBody
    Set BBX = CreateBoundingBox(part)
    
    ' ��ȡ������
    Dim Extract As HybridShapeExtract
    Set Extract = CreateExtract(part, MBD, BBX)
    
    ' ���㼫ֵ��
    Dim extremums As Collection
    Set extremums = CalculateExtremums(part, Extract, BBX)
    
    ' �������벢��������
    Dim dimensions As Variant
    dimensions = MeasureAndCreateParameters(part, extremums, BBX)
    
    ' �����߽�򼸺���
    CreateBoundingBoxGeometry part, BBX, dimensions, extremums
    
    ' ������ʾ����
    SetDisplayProperties Doc, BBX
    
    ' �ָ�ͬ��
    CATIA.HSOSynchronized = True
    
    MsgBox "�����������������"
End Sub

' ����1����֤��ĵ�����
Function ValidateActiveDocument() As Boolean
    ValidateActiveDocument = True
    If TypeName(CATIA.ActiveDocument) <> "PartDocument" Then
        MsgBox ("��ȷ�ϵ�ǰ��ĵ�����Ҫ�����߽����ĵ�")
        ValidateActiveDocument = False
    End If
End Function

' ����2�������߽������
Function CreateBoundingBox(part As Part) As HybridBody
    Dim BBX As HybridBody
    Set BBX = part.HybridBodies.Add
    BBX.Name = "Boundingbox"
    Set CreateBoundingBox = BBX
End Function

' ����3����ȡ������
Function CreateExtract(part As Part, MBD As Body, BBX As HybridBody) As HybridShapeExtract
    Dim HSF As HybridShapeFactory
    Set HSF = part.HybridShapeFactory
    
    Dim ref As Reference
    Set ref = part.CreateReferenceFromObject(MBD)
    
    Dim Extract As HybridShapeExtract
    Set Extract = HSF.AddNewExtract(ref)
    BBX.AppendHybridShape Extract
    Extract.Compute
    
    ' ��ӵ�ѡ���Ա��������
    CATIA.ActiveDocument.Selection.Add Extract
    
    Set CreateExtract = Extract
End Function

' ����4�����㼫ֵ��
Function CalculateExtremums(part As Part, Extract As HybridShapeExtract, BBX As HybridBody) As Collection
    Dim HSF As HybridShapeFactory
    Set HSF = part.HybridShapeFactory
    
    Dim ref As Reference
    Set ref = part.CreateReferenceFromObject(Extract)
    
    ' ������������
    Dim xDir As HybridShapeDirection
    Dim yDir As HybridShapeDirection
    Dim zDir As HybridShapeDirection
    Set xDir = HSF.AddNewDirectionByCoord(1, 0, 0)
    Set yDir = HSF.AddNewDirectionByCoord(0, 1, 0)
    Set zDir = HSF.AddNewDirectionByCoord(0, 0, 1)
    
    ' ���������ļ�ֵ��
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
    
    ' ���²���
    part.Update
    
    ' �����м�ֵ��洢�ڼ����з���
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

' ����5���������벢��������
Function MeasureAndCreateParameters(part As Part, extremums As Collection, BBX As HybridBody) As Variant
    Dim Doc As Document
    Set Doc = CATIA.ActiveDocument
    
    ' ��ȡSPA����̨���ڲ���
    Dim WB As Workbench
    Set WB = Doc.GetWorkbench("SPAWorkbench")
    
    ' �������
    Dim Mes(2) As Measurable
    Dim Arr(5) As Double
    Dim DisX As Double, DisY As Double, DisZ As Double
    Dim xmaxc As Double, xminc As Double
    Dim ymaxc As Double, yminc As Double
    Dim zmaxc As Double, zminc As Double
    
    ' ����X�������
    Set Mes(0) = WB.GetMeasurable(extremums("xmax"))
    Mes(0).GetMinimumDistancePoints extremums("xmin"), Arr
    DisX = Abs(Arr(3) - Arr(0))
    xmaxc = Arr(0): xminc = Arr(3)
    
    ' ����Y�������
    Set Mes(1) = WB.GetMeasurable(extremums("ymax"))
    Mes(1).GetMinimumDistancePoints extremums("ymin"), Arr
    DisY = Abs(Arr(4) - Arr(1))
    ymaxc = Arr(1): yminc = Arr(4)
    
    ' ����Z�������
    Set Mes(2) = WB.GetMeasurable(extremums("zmax"))
    Mes(2).GetMinimumDistancePoints extremums("zmin"), Arr
    DisZ = Abs(Arr(5) - Arr(2))
    zmaxc = Arr(2): zminc = Arr(5)
    
    ' �����������ߴ�
    Dim product2 As Part
    Set product2 = Doc.Part
    Dim parameters1 As Parameters
    Set parameters1 = product2.Parameters 'UserRefProperties
    
    Dim length1 As Length, length2 As Length, length3 As Length
    Set length1 = parameters1.CreateDimension("X��", "LENGTH", DisX)
    Set length2 = parameters1.CreateDimension("Y��", "LENGTH", DisY)
    Set length3 = parameters1.CreateDimension("Z��", "LENGTH", DisZ)
    
    ' ���ز����������
    Dim results(8) As Variant
    results(0) = DisX: results(1) = DisY: results(2) = DisZ
    results(3) = xmaxc: results(4) = xminc
    results(5) = ymaxc: results(6) = yminc
    results(7) = zmaxc: results(8) = zminc
    
    MeasureAndCreateParameters = results
End Function

' ����6�������߽�򼸺���
Sub CreateBoundingBoxGeometry(part As Part, BBX As HybridBody, dimensions As Variant, extremums As Collection)
    Dim HSF As HybridShapeFactory
    Set HSF = part.HybridShapeFactory
    
    ' ��ȡ����ֵ
    Dim DisX As Double, DisY As Double, DisZ As Double
    Dim xmaxc As Double, xminc As Double
    Dim ymaxc As Double, yminc As Double
    Dim zmaxc As Double, zminc As Double
    DisX = dimensions(0): DisY = dimensions(1): DisZ = dimensions(2)
    xmaxc = dimensions(3): xminc = dimensions(4)
    ymaxc = dimensions(5): yminc = dimensions(6)
    zmaxc = dimensions(7): zminc = dimensions(8)
    
    ' ������������
    Dim xDir As HybridShapeDirection
    Dim yDir As HybridShapeDirection
    Dim zDir As HybridShapeDirection
    Set xDir = HSF.AddNewDirectionByCoord(1, 0, 0)
    Set yDir = HSF.AddNewDirectionByCoord(0, 1, 0)
    Set zDir = HSF.AddNewDirectionByCoord(0, 0, 1)
    
    ' ������һ����
    Dim p1 As HybridShapePointCoord
    Set p1 = HSF.AddNewPointCoord(xmaxc, yminc, zminc)
    BBX.AppendHybridShape p1
    p1.Compute
    CATIA.ActiveDocument.Selection.Add p1
    
    ' �����ڶ�����
    Dim p2 As HybridShapePointCoord
    Set p2 = HSF.AddNewPointCoord(xminc, yminc, zminc)
    BBX.AppendHybridShape p2
    p2.Compute
    CATIA.ActiveDocument.Selection.Add p2
    
    ' �����߶�
    Dim ln As HybridShapeLinePtPt
    Set ln = HSF.AddNewLinePtPt(p1, p2)
    BBX.AppendHybridShape ln
    ln.Compute
    CATIA.ActiveDocument.Selection.Add ln
    
    ' ��Y��������
    Dim ext As HybridShapeExtrude
    Set ext = HSF.AddNewExtrude(ln, DisY, 0, yDir)
    BBX.AppendHybridShape ext
    ext.Compute
    CATIA.ActiveDocument.Selection.Add ext
    
    ' ��������߽�
    Dim bound As HybridShapeBoundary
    Set bound = HSF.AddNewBoundaryOfSurface(ext)
    BBX.AppendHybridShape bound
    bound.Compute
    CATIA.ActiveDocument.Selection.Add bound
    
    ' ��Z��������
    Dim ext2 As HybridShapeExtrude
    Set ext2 = HSF.AddNewExtrude(bound, DisZ, 0, zDir)
    BBX.AppendHybridShape ext2
    ext2.Compute
    CATIA.ActiveDocument.Selection.Add ext2
    
    ' ƽ�Ʋ���
    Dim trans As HybridShapeTransform
    Set trans = HSF.AddNewTranslate(ext, zDir, DisZ)
    BBX.AppendHybridShape trans
    trans.Compute
    CATIA.ActiveDocument.Selection.Add trans
    
    ' �ϲ�����
    Dim asm As HybridShapeAssemble
    Set asm = HSF.AddNewJoin(ext, ext2)
    asm.AddElement trans
    BBX.AppendHybridShape asm
    asm.Compute
    
    ' �������ݼ�
    Dim eles As Variant
    eles = HSF.AddNewDatums(asm)
    BBX.AppendHybridShape eles(0)
    eles(0).Name = "Bounding box of " & part.Bodies.Item(1).Name
    HSF.DeleteObjectForDatum asm
    
    ' ����ѡ��
    CATIA.ActiveDocument.Selection.Delete
    CATIA.ActiveDocument.Selection.Add eles(0)
End Sub

' ����7��������ʾ����
Sub SetDisplayProperties(Doc As Document, BBX As HybridBody)
    ' ����͸����
    Doc.Selection.VisProperties.SetRealOpacity 100, 1
    Doc.Selection.Clear
End Sub