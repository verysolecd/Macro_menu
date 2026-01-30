Attribute VB_Name = "MDL_Shapeinfo"
Private mdict, HSF, oSel, oParas, spa
Private Const mdlname As String = "MDL_Shapeinfo"
Sub tube()
'--判断是否是part，当前代码只能运行在part中，修改后才能在总成中运行，例如增加遍历
If TypeName(CATIA.ActiveDocument) <> "PartDocument" Then
'        Exit Sub
End If
'初始化各类参数
    Set mdict = InitDic
        Set odoc = CATIA.ActiveDocument
        Set spa = odoc.GetWorkbench("SPAWorkbench")
        Set oPrt = odoc.part
        Set oSel = odoc.Selection
        Set oParas = oPrt.Parameters
        Set HSF = odoc.part.HybridShapeFactory
        Set HBS = odoc.part.HybridBodies
'获取所有的shapes
Set oHb = oPrt ' HBS.item(1)
Call GetShapesByRecursion(oPrt)
'Call GetShapesByParameters(oprt)
'将获取的shape增加到list
    Set lst = Initlst
    For Each key In mdict.keys
        Set itube = mdict(key)
        lst.Add itube
    Next
'对lst 产品分类处理
For Each lstitm In lst
    Select Case LCase(TypeName(lstitm))
    Case LCase("HybridShapeInstance")  '这里是实例化的接头类，对应零件内UDF实例
    '以下代码获取实例的参数信息，使用dict输出，也可以用其他输出方式
         Dim ipa: Set ipa = InitDic
            ipa.Add lstitm.Name, "我是" & strbflast(lstitm.Name, ".")
            For i = 1 To 10  '如果实例有更多参数，10增加
            On Error Resume Next
                Set pa = lstitm.GetParameterFromPosition(i)
                If Not pa Is Nothing Then
                nameInRel = GetInternalName(pa)
                'ipa.Add itm.Name, 333333
                ipa.Add nameInRel, pa.value
            End If
            Err.Clear
            On Error GoTo 0
            Next
           showdict ipa
            'Debug.Print lstitm.Name & "_直径是" & "XX" & "_厚度是" & "XX"
    Case LCase("ThickSurface")  '判定为是加厚曲面
        '以下代码获取加厚曲面的参数信息，使用了debug.print输出，可以修改为输出到excel
        '扩展代码以输出更多信息
         ' Debug.Print lstitm.Name & "是加厚曲面"
             tk = lstitm.TopOffset.value '获取加厚曲面的厚度
                Set oSweep = GetParentSweep(oPrt, lstitm) '获取加厚曲面的sweep
                    Set oCurve = GetParentcurve(oPrt, oSweep) '获取的sweep的父级曲线
                    lg = getlength(oCurve)   '获取曲线长度
                Debug.Print lstitm.Name & "_长度是"; Round(lg, 1) & "_厚度是" & tk  '输出厚度和长度，
    End Select
 Next
End Sub
Sub GetShapesByParameters(oPrt)
    Set paras = oPrt.Parameters
    For Each P In paras
        On Error Resume Next
        Dim parentObj
        Set parentObj = P.Parent
        If Not parentObj Is Nothing Then
            If TypeName(parentObj) <> "Parameters" Then
                If HSF.GetGeometricalFeatureType(parentObj) = 7 Then
                    Dim oname As String
                    oname = GetInternalName(parentObj)
                    If Not mdict.Exists(oname) Then
                        mdict.Add oname, parentObj
                        ' Debug.Print "Found by Param: " & oname
                    End If
                End If
            End If
        End If
        Err.Clear
    Next
    On Error GoTo 0
End Sub
Sub recurallBody(iHB, HSF)
    Dim Shps: Set Shps = iHB.HybridShapes
    For Each shp In Shps
        If HSF.GetGeometricalFeatureType(shp) = 7 Then
            oname = GetInternalName(shp)
             If mdict.Exists(oname) = False Then
                Set mdict(oname) = shp
                ' Debug.Print TypeName(Shp) & " : " & oname & " : " & Shp.Name
            End If
         End If
    Next
    ' Recursively process child HybridBodies
    If iHB.HybridBodies.count > 0 Then
            For Each chb In iHB.HybridBodies
                Call recurallBody(chb, HSF)
             Next
    End If
End Sub
Sub recurAyo(ayo)
    Dim colls: Set itm = ayo.Products
    For Each itm In colls
        Call recurFunc(itm)
    Next
    If ayo.Products.count > 0 Then
            For Each ctm In ayo.Products
                Call recurAyo(ctm)
             Next
    End If
End Sub
Function recurFunc(itm)
  'XCXX
End Function
Sub GetShapesByRecursion(iHB)
CATIA.RefreshDisplay = False
    On Error Resume Next
      Set Shps = iHB.HybridShapes
        If Not Shps Is Nothing Then
            For Each shp In Shps
               iType = HSF.GetGeometricalFeatureType(shp)
                If iType = 7 Then
                    internalName = GetInternalName(shp)
                    If Not mdict.Exists(internalName) Then
                        oSel.Clear: oSel.Add shp
                        Set realShp = oSel.item(1).value: oSel.Clear
                        mdict.Add internalName, realShp
                    End If
                End If
            Next
        End If
    Err.Clear
    If iHB.HybridBodies.count > 0 Then
        For Each childHB In iHB.HybridBodies
            Call GetShapesByRecursion(childHB)
        Next
    End If
     Err.Clear
    On Error GoTo 0
    CATIA.RefreshDisplay = True
End Sub
Function GetParentSweep(targetPart, thickSurf) As Object
    On Error Resume Next
    Set GetParentSweep = Nothing
    Set inputRef = thickSurf.Surface
    If inputRef Is Nothing Then Exit Function
    refName = inputRef.DisplayName
    Dim nameParts() As String
    nameParts = Split(refName, "/")
    Dim i As Integer
    Dim potentialName As String
    Dim foundObj As Object
    For i = UBound(nameParts) To 0 Step -1
        potentialName = nameParts(i)
        If InStr(potentialName, "Face") = 0 And InStr(potentialName, "Edge") = 0 And InStr(potentialName, "Vertex") = 0 Then
            Err.Clear
            Set foundObj = targetPart.FindObjectByName(potentialName)
            If Err.Number = 0 And Not foundObj Is Nothing Then
'                If InStr(TypeName(foundObj), "HybridShapeSweep") > 0 Or TypeName(foundObj) = "HybridShapeSweep" Then
                    Set GetParentSweep = foundObj
                    Exit Function
'                End If
            End If
        End If
    Next i
    On Error GoTo 0
End Function
Function GetParentcurve(targetPart, thickSurf) As Object
    On Error Resume Next
    Set GetParentcurve = Nothing
    Set inputRef = thickSurf.FirstGuideCrv
    If inputRef Is Nothing Then Exit Function
    refName = inputRef.DisplayName
    Dim nameParts() As String
    nameParts = Split(refName, "/")
    Dim i As Integer
    Dim potentialName As String
    Dim foundObj As Object
    For i = UBound(nameParts) To 0 Step -1
        potentialName = nameParts(i)
        If InStr(potentialName, "Face") = 0 And InStr(potentialName, "Edge") = 0 And InStr(potentialName, "Vertex") = 0 Then
            Err.Clear
            Set foundObj = targetPart.FindObjectByName(potentialName)
            If Err.Number = 0 And Not foundObj Is Nothing Then
'                If InStr(TypeName(foundObj), "HybridShapeSweep") > 0 Or TypeName(foundObj) = "HybridShapeSweep" Then
                    Set GetParentcurve = foundObj
                    Exit Function
'                End If
            End If
        End If
    Next i
    On Error GoTo 0
End Function
Function getlength(itm)
 getlength = 0
   If Not itm Is Nothing Then
        Set meas = spa.GetMeasurable(itm)
        getlength = meas.Length
    End If
End Function
Function InitDic(Optional compareMode As Long = vbBinaryCompare) As Object
    Dim Dic As Object
    Set Dic = CreateObject("Scripting.Dictionary")
    Dic.compareMode = compareMode
    Set InitDic = Dic
End Function
Function Initlst() As Object
    Set Initlst = CreateObject("System.Collections.ArrayList")
End Function
Function GetInternalName$(aoj)
    If IsNothing(aoj) Then
        GetInternalName = Empty: Exit Function
    End If
    GetInternalName = aoj.GetItem("ModelElement").internalName
End Function

Function strbflast(str, iext)
Dim idx
idx = InStrRev(str, iext)
If idx > 0 Then
        strbflast = Left(str, idx - 1)
    Else
        strbflast = str
    End If
End Function
