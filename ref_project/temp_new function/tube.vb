Private mdict
Sub tube()
    Set mdict = KCL.InitDic
' If Not CanExecute("PartDocument") Then
'        Exit Sub
'    End If
    Dim oDoc: Set oDoc = CATIA.ActiveDocument
    Set oprt = oDoc.part
    Dim HSF:  Set HSF = oDoc.part.HybridShapeFactory
    Dim HBS: Set HBS = oDoc.part.HybridBodies
    Dim oSel: Set oSel = oDoc.Selection
    Dim otube
  '    Set otube = oshapes.item(1)
'    Set oface = otube.Surface
'    Set ocr = oshapes.item(4)
    Set Shps = HBS.item(2).HybridShapes
    Set paras = oprt.Parameters
         For Each Shp In paras
           If TypeName(Shp.Parent) <> "Parameters" Then
                If HSF.GetGeometricalFeatureType(Shp.Parent) = 7 Then
                    oname = KCL.GetInternalName(Shp.Parent)
                    If mdict.Exists(oname) = False Then
                            Set mdict(oname) = Shp.Parent
                              Debug.Print TypeName(Shp.Parent) & "！！！！！！" & oname & Shp.Name
                    End If
                End If
            End If
        Next
'Next
'
Set lst = KCL.InitLst
    For Each key In mdict.keys
        Set itube = mdict(key)
'        Debug.Print itube.Name
        lst.Add itube
    Next
MsgBox "zahntong "
End Sub
Sub recurallBody(ihb)
    Shps = ihb.Shapes
    For Each Shp In Shps
        If HSF.GetGeometricalFeatureType(Shp) = 7 Then
            oname = KCL.GetInternalName(Shp)
             If mdict.Exists(oname) = False Then
                Set mdict(oname) = Shp
            End If
         End If
    Next
    If ihb.HybridBodies.count > 0 Then
            For Each HB In ihb.HybridBodies
                Call recurallBody(HB)
             Next
    End If
End Sub