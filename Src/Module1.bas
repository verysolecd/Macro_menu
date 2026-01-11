Attribute VB_Name = "Module1"
Private mdict, HSF, osel, oParas, spa
Sub tube()
' If Not CanExecute("PartDocument") Then
'        Exit Sub
'    End If
    Set mdict = KCL.InitDic
        Set oDoc = CATIA.ActiveDocument
        Set spa = oDoc.GetWorkbench("SPAWorkbench")
        Set oprt = oDoc.part
        Set osel = oDoc.Selection
        Set oParas = oprt.Parameters
        Set HSF = oDoc.part.HybridShapeFactory
        Set HBS = oDoc.part.HybridBodies
    
Set oHb = oprt ' HBS.item(1)
Call GetShapesByRecursion(oprt)
'Call GetShapesByParameters(oprt)
    Set lst = KCL.InitLst
    For Each key In mdict.keys
        Set itube = mdict(key)
        lst.Add itube
    Next
    
For Each lstitm In lst
 
    '对lst 产品分类
    
    Select Case LCase(TypeName(lstitm))
    
    Case LCase("HybridShapeInstance")
          Debug.Print lstitm.Name & "_直径是" & 9&; "__厚度是" & 9
          
    On Error Resume Next
        For i = 1 To 10
            Set pa = itm.GetParameterFromPosition(i)
            If Not pa Is Nothing Then
            nameInRel = KCL.GetInternalName(pa)
            ipa.Add itm.Name, 333333
            ipa.Add nameInRel, pa.value
        End If
        Next
    Case LCase("ThickSurface")
         ' Debug.Print lstitm.Name & "是加厚曲面"
          '获取加厚曲面的厚度
             tk = lstitm.TopOffset.value
          
            '获取加厚曲面的sweep
                Set oSweep = GetParentSweep(oprt, lstitm)
           '获取的sweep的父级曲线
                    Set oCurve = GetParentcurve(oprt, oSweep)
            '获取曲线长度
                    lg = getlength(oCurve)
                Debug.Print lstitm.Name & "___长度是"; Round(lg, 1) & "__厚度是" & tk
    End Select
 Next
 
 
' osel.Clear
' osel.Add itm
' Set itm = osel.item(1).value
 
 
 
'获取实例的参数数组和直径
    
    KCL.showdict ipa
    
  
    
    MsgBox "iue"
End Sub
Sub GetShapesByParameters(oprt)
    Set paras = oprt.Parameters
    For Each P In paras
        On Error Resume Next
        Dim parentObj
        Set parentObj = P.Parent
        If Not parentObj Is Nothing Then
            If TypeName(parentObj) <> "Parameters" Then
                If HSF.GetGeometricalFeatureType(parentObj) = 7 Then
                    Dim oname As String
                    oname = KCL.GetInternalName(parentObj)
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
            oname = KCL.GetInternalName(shp)
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
    On Error Resume Next
      Set Shps = iHB.HybridShapes
        If Not Shps Is Nothing Then
            For Each shp In Shps
               iType = HSF.GetGeometricalFeatureType(shp)
                If iType = 7 Then
                    internalName = KCL.GetInternalName(shp)
                    If Not mdict.Exists(internalName) Then
                        osel.Clear: osel.Add shp
                        Set realShp = osel.item(1).value: osel.Clear
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


