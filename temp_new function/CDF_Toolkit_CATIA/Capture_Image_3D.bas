Attribute VB_Name = "Capture_Image_3D"
Option Explicit

Sub CATMain()

IntCATIA
If TypeName(oActDoc) <> "ProductDocument" Then
   If TypeName(oActDoc) <> "PartDocument" Then
        MsgBox "当前文件不是CATIA零件或产品 !" & vbCrLf & "此命令只能在零件或产品文件中运行!"
        Exit Sub
   End If
End If
On Error Resume Next
'-----------设置设计模式-------------
CATIA.DisplayFileAlerts = False
'Dim settingControllers1
'Dim cacheSettingAtt1
'Set settingControllers1 = CATIA.SettingControllers
'Set cacheSettingAtt1 = settingControllers1.Item("CATSysCacheSettingCtrl")
'cacheSettingAtt1.ActivationMode = 1
'Set cacheSettingAtt1 = Nothing
'Set settingControllers1 = Nothing
Dim oProdxx
Set oProdxx = oActDoc.Product
oProdxx.ApplyWorkMode DESIGN_MODE
'-----------设置设计模式-------------
On Error Resume Next
'hideConstraints oActDoc.Product
'HideDatums oActDoc.Product   '隐藏基准面和坐标系
HideDatumsConstraints oActDoc, 1
Dim OutPutFolder

OutPutFolder = oCATVBA_Folder("Temp")

CaptureChildrenImage oProdxx, OutPutFolder
On Error Resume Next
If MsgBox("操作完成!" & _
        vbCrLf & _
        vbCrLf & "截图放在 " & OutPutFolder & " 下", vbOKCancel) = vbOK Then
    Shell "explorer.exe " & OutPutFolder, vbNormalFocus
End If

End Sub
Sub CaptureChildrenImage(oProduct9, OutputFolder9)
CATIA.RefreshDisplay = False
On Error Resume Next
Dim myViewer1
Dim color(2)
Set myViewer1 = CATIA.ActiveWindow.ActiveViewer

myViewer1.GetBackgroundColor color
myViewer1.PutBackgroundColor Array(1, 1, 1)
CATIA.ActiveWindow.Layout = 1

Dim specsAndGeomWindow1 As SpecsAndGeomWindow
Set specsAndGeomWindow1 = CATIA.ActiveWindow
Dim viewer3D1 As Viewer3D
Set viewer3D1 = specsAndGeomWindow1.ActiveViewer
    viewer3D1.RenderingMode = 1 ' "catRenderShadingWithEdges"
    viewer3D1.Reframe
    viewer3D1.CaptureToFile 5, OutputFolder9 & "\" & Replace(oProduct9.ReferenceProduct.Parent.Name, ".", "_") & ".jpg"
If oProduct9.Products.Count = 0 Then
CATIA.ActiveWindow.Layout = 2 ' catWindowSpecsAndGeom
myViewer1.PutBackgroundColor color
' This VBA Macro Developed by Charles.Tang
' WeChat Chtang80,CopyRight reserved
Exit Sub
End If

Dim selection9 As Selection
Set selection9 = CATIA.ActiveDocument.Selection
selection9.Clear
Dim visPropertySet9 As VisPropertyset
Set visPropertySet9 = selection9.VisProperties
Dim oProducts9
Set oProducts9 = oProduct9.Products

Dim i
For i = 1 To oProducts9.Count
selection9.Add oProducts9.Item(i)
Next
visPropertySet9.SetShow 1
selection9.Clear

For i = 1 To oProducts9.Count
selection9.Add oProducts9.Item(i)
visPropertySet9.SetShow 0
selection9.Clear
Call CaptureChildrenImage(oProducts9.Item(i), OutputFolder9)
selection9.Add oProducts9.Item(i)
visPropertySet9.SetShow 1
selection9.Clear
Next


CATIA.ActiveWindow.Layout = 2 ' catWindowSpecsAndGeom
myViewer1.PutBackgroundColor color

'show all components

For i = 1 To oProducts9.Count
selection9.Add oProducts9.Item(i)
Next

visPropertySet9.SetShow 0

selection9.Clear
On Error GoTo 0

Set myViewer1 = Nothing
Set specsAndGeomWindow1 = Nothing
Set viewer3D1 = Nothing
Set selection9 = Nothing
Set visPropertySet9 = Nothing
Set oProducts9 = Nothing

'CATIA.RefreshDisplay = True
End Sub

'Sub HideDatums(oProducti As Product)
'If oProducti.Products.Count = 0 Then
'
'HidePartDatums oProducti.ReferenceProduct.Parent.Part
'Else
'Dim j
'    For j = 1 To oProducti.Products.Count
'    Call HideDatums(oProducti.Products.Item(j))
'    Next
'End If
'End Sub
'Sub HidePartDatums(oParti)
'Dim pSel As Selection
'Set pSel = oParti.Parent.Selection
'pSel.Clear
'
'Dim colAxSys, OrigEl
'Set colAxSys = oParti.AxisSystems
'Set OrigEl = oParti.OriginElements
'
'Dim i As Integer
'For i = 1 To colAxSys.Count
'pSel.Add colAxSys.Item(i)
'Next
'
'pSel.Add OrigEl.PlaneXY
'pSel.Add OrigEl.PlaneZX
'pSel.Add OrigEl.PlaneYZ
'' This VBA Macro Developed by Charles.Tang
'' WeChat Chtang80,CopyRight reserved
'
'Dim hybridBodies1 As Object
'Set hybridBodies1 = oParti.HybridBodies
'
'For i = 1 To hybridBodies1.Count
'pSel.Add hybridBodies1.Item(i)
'Next
'
'Dim bd As Object
'Dim sks As Object
'
'For i = 1 To oParti.Bodies.Count
'     Dim j As Integer
'     Set sks = oParti.Bodies.Item(i).Sketches
'     For j = 1 To sks.Count
'        pSel.Add sks.Item(j)
'     Next
'Next
'
'
'
'pSel.VisProperties.SetShow 1
'
'pSel.Clear
'
'End Sub
'Sub hideConstraints(oProd As Object)
'Dim pSel As Selection
'Set pSel = CATIA.ActiveDocument.Selection
'pSel.Clear
'
'If oProd.Products.Count = 0 Then
'Exit Sub
'End If
'
'Dim oConstraints As Object
'Dim i2 As Integer
'Set oConstraints = oProd.Connections("CATIAConstraints")
'For i2 = 1 To oConstraints.Count
'pSel.Add oConstraints.Item(i2)
''Debug.Print oConstraints.Item(i2).Name & " hided"
'Next
'pSel.VisProperties.SetShow 1
'
'Dim oProdi As Object
'
'For Each oProdi In oProd.Products
''Debug.Print "处理产品" & oProdi.Name
'hideConstraints oProdi
'Next
'
'End Sub
Sub HideDatumsConstraints(prodoc As Document, intshow As Integer)

On Error Resume Next
Dim selection1 As Selection
Set selection1 = prodoc.Selection
selection1.Clear

selection1.Search "(((CATStFreeStyleSearch.Plane + CATPrtSearch.Plane) + CATGmoSearch.Plane) + CATSpdSearch.Plane),all"

selection1.VisProperties.SetShow intshow

selection1.Clear

selection1.Search "(((CATStFreeStyleSearch.AxisSystem + CATPrtSearch.AxisSystem) + CATGmoSearch.AxisSystem) + CATSpdSearch.AxisSystem),all"


selection1.VisProperties.SetShow intshow

selection1.Clear

selection1.Search "((((((CATStFreeStyleSearch.Point + CAT2DLSearch.2DPoint) + CATSketchSearch.2DPoint) + CATDrwSearch.2DPoint) + CATPrtSearch.Point) + CATGmoSearch.Point) + CATSpdSearch.Point),all"
selection1.VisProperties.SetShow intshow

selection1.Clear

selection1.Search "((((((CATStFreeStyleSearch.Curve + CAT2DLSearch.2DCurve) + CATSketchSearch.2DCurve) + CATDrwSearch.2DCurve) + CATPrtSearch.Curve) + CATGmoSearch.Curve) + CATSpdSearch.Curve),all"
selection1.VisProperties.SetShow intshow

selection1.Clear

selection1.Search "(((CATStFreeStyleSearch.Surface + CATPrtSearch.Surface) + CATGmoSearch.Surface) + CATSpdSearch.Surface),all"
selection1.VisProperties.SetShow intshow

selection1.Clear

selection1.Search "(((((((CATProductSearch.MfConstraint + CATStFreeStyleSearch.MfConstraint) + CATAsmSearch.MfConstraint) + CAT2DLSearch.MfConstraint) + CATSketchSearch.MfConstraint) + CATDrwSearch.MfConstraint) + CATPrtSearch.MfConstraint) + CATSpdSearch.MfConstraint),all"
selection1.VisProperties.SetShow intshow
selection1.Clear
Err.Clear
End Sub
