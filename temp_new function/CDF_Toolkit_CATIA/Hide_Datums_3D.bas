Attribute VB_Name = "Hide_Datums_3D"
Option Explicit
#If VBA7 Then
Private Declare PtrSafe Function GetKeyState Lib "user32" (ByVal nVirtKey As Long) As Integer
#Else
Private Declare Function GetKeyState Lib "user32" (ByVal nVirtKey As Long) As Integer
#End If

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
Dim oProdxx
Set oProdxx = oActDoc.Product
oProdxx.ApplyWorkMode DESIGN_MODE
'-----------设置设计模式-------------

If (GetKeyState(vbKeyShift) And &H8000&) Then
    hideConstraints oActDoc.Product, False
    HideDatums oActDoc.Product, False
Else
    hideConstraints oActDoc.Product
    HideDatums oActDoc.Product
End If
oProdxx.ApplyWorkMode DEFAULT_MODE
Set oProdxx = Nothing
End Sub
Sub HideDatums(oProducti As Product, Optional h As Boolean = True)
oProducti.ApplyWorkMode DESIGN_MODE
If oProducti.Products.Count = 0 Then

HidePartDatums oProducti.ReferenceProduct.Parent.Part, h
Else
Dim j
    For j = 1 To oProducti.Products.Count
    Call HideDatums(oProducti.Products.Item(j), h)
    Next
End If
oProducti.ApplyWorkMode DEFAULT_MODE
End Sub
Sub HidePartDatums(oParti, Optional h As Boolean = True)
Dim pSel As Selection
Set pSel = oParti.Parent.Selection
pSel.Clear

Dim colAxSys, OrigEl
Set colAxSys = oParti.AxisSystems
Set OrigEl = oParti.OriginElements

Dim i As Integer
For i = 1 To colAxSys.Count
pSel.Add colAxSys.Item(i)
Next

pSel.Add OrigEl.PlaneXY
pSel.Add OrigEl.PlaneZX
pSel.Add OrigEl.PlaneYZ
' This VBA Macro Developed by Charles.Tang
' WeChat Chtang80,CopyRight reserved
Dim hybridBodies1 As Object
Set hybridBodies1 = oParti.HybridBodies

For i = 1 To hybridBodies1.Count
pSel.Add hybridBodies1.Item(i)
Next

Dim bd As Object
Dim sks As Object

For i = 1 To oParti.Bodies.Count
     Dim j As Integer
     Set sks = oParti.Bodies.Item(i).Sketches
     For j = 1 To sks.Count
        pSel.Add sks.Item(j)
     Next
Next
If h = True Then
pSel.VisProperties.SetShow 1
Else
pSel.VisProperties.SetShow 0
End If

pSel.Clear

End Sub
Sub hideConstraints(oProd As Object, Optional h As Boolean = True)
Dim pSel As Selection
Set pSel = CATIA.ActiveDocument.Selection
pSel.Clear

If oProd.Products.Count = 0 Then
Exit Sub
End If

Dim oConstraints As Object
Dim i2 As Integer
Set oConstraints = oProd.Connections("CATIAConstraints")
For i2 = 1 To oConstraints.Count
pSel.Add oConstraints.Item(i2)
'Debug.Print oConstraints.Item(i2).Name & " hided"
Next
If h = True Then
pSel.VisProperties.SetShow 1
Else
pSel.VisProperties.SetShow 0
End If

Dim oProdi As Object

For Each oProdi In oProd.Products
'Debug.Print "处理产品" & oProdi.Name
hideConstraints oProdi, h
Next

End Sub
