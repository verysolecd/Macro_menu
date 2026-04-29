Public selection1 As Selection, selection2 As Selection
Sub CATMain()
 On Error Resume Next
	oPath = UserForm1.Path.Text
	EbomName = UserForm1.TextBox6.Text
Set xlApp = CreateObject("Excel.Application")
	xlApp.Caption = "明细"
	xlApp.Workbooks.Open (EbomName)
    If Err.Number = 0 Then
'    xlApp.Visible = True
    xlApp.ReferenceStyle = xlR1C1
    Else
    Exit Sub
    End If
Set xlSheet1 = xlApp.Workbooks.Item(1).Sheets(1)
CATIA.DisplayFileAlerts = False
Dim documents1 As Documents
Set documents1 = CATIA.Documents
Dim productDocument1 As ProductDocument
Set productDocument1 = documents1.Add("Product")
Dim product1 As Product
Set product1 = productDocument1.Product


product1.PartNumber = xlSheet1.Cells(2, 3)
product1.Definition = xlSheet1.Cells(2, 4)
product1.Nomenclature = xlSheet1.Cells(2, 5)
product1.Update


Dim i%, iCount%, PrtName As String, PrtName_p As String, lev0%, lev1 As Range, Rng As Range, TRng As Ranges
iCount = xlApp.CountA(xlSheet1.Columns(1))
For i = 3 To iCount
    PrtName = xlSheet1.Cells(i, 3)
    Set lev1 = xlSheet1.Cells(i, 2)
        lev0 = lev1 - 1
        For R = i To 2 Step -1
            If xlSheet1.Cells(R, 2) = lev0 Then
            Rn = xlSheet1.Cells(R, 2).Row
            Exit For
            End If
        Next
        PrtName_p = xlSheet1.Cells(Rn, 3)
    Dim products1 As Products
    Set products1 = documents1.Item(PrtName_p & ".CATProduct").Product.Products
    If xlSheet1.Cells(i, 14) = "CATProduct" Then
        ProductName = oPath & "\" & PrtName & ".CATProduct"
        If Len(Dir(ProductName)) > 0 Then
            Dim arrayOfVariantOfBSTR1(0)
            arrayOfVariantOfBSTR1(0) = ProductName
            Set products1Variant = products1
            products1Variant.AddComponentsFromFiles arrayOfVariantOfBSTR1, "All"
            Set products3 = documents1.Item(PrtName & ".CATProduct").Product.Products
            i = i + products3.Count
        Else
            Set part2 = products1.AddNewComponent("Product", PrtName)
            part2.Definition = xlSheet1.Cells(i, 4)
            part2.Nomenclature = xlSheet1.Cells(i, 5)
            part2.Revision = xlSheet1.Cells(i, 10)
            part2.DescriptionRef = xlSheet1.Cells(i, 19)
            part2.Update    '更新文件
            Set productDocument2 = CATIA.Documents.Item(PrtName & ".CATProduct")
            productDocument2.SaveAs oPath & "\" & PrtName
        End If
    ElseIf xlSheet1.Cells(i, 14) = "CATPart" Then
        PartName = oPath & "\" & PrtName & ".CATPart"
        If Len(Dir(PartName)) > 0 Then
            Dim arrayOfVariantOfBSTR2(0)
            arrayOfVariantOfBSTR2(0) = PartName
            Set products4Variant = products1
            products4Variant.AddComponentsFromFiles arrayOfVariantOfBSTR2, "All"
            Set part2 = products4Variant.Item(PrtName & ".1")
        Else
            Set part2 = products1.AddNewComponent("Part", PrtName)
            part2.Definition = xlSheet1.Cells(i, 4)
            part2.Nomenclature = xlSheet1.Cells(i, 5)
            part2.Revision = xlSheet1.Cells(i, 10)
            part2.DescriptionRef = xlSheet1.Cells(i, 19)
            If InStr(PrtName, "W") = 0 Then
                Set partDocument1 = CATIA.Documents.Item(PrtName & ".CATPart")
                Set part1 = partDocument1.part
                Set Product01 = partDocument1.Product
                Mtl001 = xlSheet1.Cells(i, 7)
                Set parameters1 = part1.Parameters
                Set strParam1 = parameters1.CreateString("Material", Mtl001)
                Set parameters2 = Product01.UserRefProperties
                Set strParam1 = parameters2.CreateString("Material", "")
                Set relations1 = Product01.Relations
                Set formula1 = relations1.CreateFormula("公式.9", "", strParam1, PrtName & "\Material")
                If InStr(PrtName, "Q") = 0 And InStr(PrtName, "J") = 0 And InStr(PrtName, "S") = 0 Then
                    Tk01 = xlSheet1.Cells(i, 8)
                    Set parameters3 = part1.Parameters
                    Set length2 = parameters3.CreateDimension("Thickness", "LENGTH", Tk01)
                    Set parameters4 = Product01.UserRefProperties
                    Set length1 = parameters4.CreateDimension("Thickness", "LENGTH", 0#)
                    Set relations1 = Product01.Relations
                    Set formula1 = relations1.CreateFormula("公式.5", "", length1, PrtName & "\Thickness")
                End If
                Set parameters1 = part1.Parameters
                Set dimension1 = parameters1.CreateDimension("Density", "DENSITY", 7860#)
                Set dimension1 = parameters1.CreateDimension("Volume", "VOLUME", 0#)
                Set relations1 = part1.Relations
                Set formula1 = relations1.CreateFormula("公式.1", "", dimension1, "smartVolume(`零件几何体` )")
                Set dimension1 = parameters1.CreateDimension("Weight", "MASS", 0#)
                Set relations1 = part1.Relations
                Set formula1 = relations1.CreateFormula("公式.3", "", dimension1, "Volume *Density ")
                Set parameters1 = Product01.UserRefProperties
                Set dimension1 = parameters1.CreateDimension("Weight", "MASS", 0#)
                Set relations1 = Product01.Relations
                Set formula1 = relations1.CreateFormula("公式.7", "", dimension1, PrtName & "\Weight")
                Set hybridBodies1 = part1.HybridBodies
                Set hybridBody1 = hybridBodies1.Add()
                hybridBody1.Name = "Information"
                Set hybridBodies2 = hybridBody1.HybridBodies
                Set hybridBody2 = hybridBodies2.Add()
                hybridBody2.Name = "Boundary_box"
                Set hybridBody3 = hybridBodies2.Add()
                hybridBody3.Name = "Material_direction"
                Set hybridBody4 = hybridBodies2.Add()
                hybridBody4.Name = "Tooling_direction"
                Set hybridBody5 = hybridBodies2.Add()
                hybridBody5.Name = "GD&T"
                Set hybridBody6 = hybridBodies1.Add()
                hybridBody6.Name = "Input_data"
                Set hybridBodies3 = hybridBody6.HybridBodies
                Set hybridBody7 = hybridBodies3.Add()
                hybridBody7.Name = "Reference"
                Set hybridBody8 = hybridBodies3.Add()
                hybridBody8.Name = "Styling"
                Set hybridBody9 = hybridBodies3.Add()
                hybridBody9.Name = "Sections"
                Set hybridBody10 = hybridBodies1.Add()
                hybridBody10.Name = "Part_definition"
                Set hybridBodies4 = hybridBody10.HybridBodies
                Set hybridBody11 = hybridBodies4.Add()
                hybridBody11.Name = "Basic_surface"
                Set hybridBody12 = hybridBodies4.Add()
                hybridBody12.Name = "Flanges"
                Set hybridBody13 = hybridBodies4.Add()
                hybridBody13.Name = "Depressions"
                Set hybridBody14 = hybridBodies4.Add()
                hybridBody14.Name = "Cut"
                Set hybridBody15 = hybridBodies4.Add()
                hybridBody15.Name = "Cutouts"
                Set hybridBody16 = hybridBodies4.Add()
                hybridBody16.Name = "Holes"
                Set hybridBody17 = hybridBodies4.Add()
                hybridBody17.Name = "Ribs/Rippen"
                Set hybridBody18 = hybridBodies1.Add()
                hybridBody18.Name = "Adapter"
                Set hybridBodies5 = hybridBody18.HybridBodies
                Set hybridBody19 = hybridBodies5.Add()
                hybridBody19.Name = "Unfillet_final"
                Set hybridBody20 = hybridBodies1.Add()
                hybridBody20.Name = "Final"
            End If
            part2.Update    '更新文件
            Set productDocument2 = CATIA.Documents.Item(PrtName & ".CATPart")
            productDocument2.SaveAs oPath & "\" & PrtName
        End If
    End If
    If xlSheet1.Cells(i, 6) > 1 Then
    n = xlSheet1.Cells(i, 6)
        For i2 = 1 To n - 1
            Set productDocument1i = CATIA.ActiveDocument
            Set selection1 = productDocument1i.Selection
            selection1.Clear
            selection1.Add part2
            selection1.Copy
            Set selection2 = productDocument1i.Selection
            selection2.Clear
            selection2.Add products1
            selection2.Paste
        Next
    End If
    Set products1 = Nothing
    Set products3 = Nothing
    Set part2 = Nothing
Next
productDocument1.SaveAs oPath & "\" & CATIA.ActiveDocument.Name '& ".CATProduct"    '保存文件
xlApp.Workbooks.Close
CATIA.DisplayFileAlerts = True
End Sub