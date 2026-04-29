Attribute VB_Name = "Body_ShowAll_3D"
Option Explicit
Public alreadysh    '如果一个零件被多次装配，反显时只需要操作一次part的几何体
Sub CATMain()
Set alreadysh = CreateObject("Scripting.Dictionary") '存储已经操作过的零件，key为料号，字典
On Error Resume Next
IntCATIA
If TypeName(oActDoc) = "ProductDocument" Then
ShowallBodies oActDoc.Product
ElseIf TypeName(oActDoc) = "PartDocument" Then
ShowallBodies2 oActDoc.Part
Else
        MsgBox "当前文件不是CATIA零件或产品 !" & vbCrLf & "此命令只能在零件或产品文件中运行!"
        Exit Sub
End If
alreadysh.RemoveAll
End Sub
Public Sub ShowallBodies(oProduct, Optional reverse As Boolean = False)
oProduct.ApplyWorkMode DESIGN_MODE
Set oSel = oActDoc.Selection
oSel.Clear
oSel.Add oProduct
oSel.VisProperties.SetShow 0 '0show,1noshow
oSel.Clear

Dim j
For j = 1 To oProduct.Products.Count
On Error Resume Next
If oProduct.Products.Item(j).Products.Count = 0 Then

    oProduct.Products.Item(j).ApplyWorkMode DESIGN_MODE
    oSel.Add oProduct.Products.Item(j)
    oSel.VisProperties.SetShow 0
    If reverse = False Then
    ShowallBodies2 oProduct.Products.Item(j).ReferenceProduct.Parent.Part, reverse
    ElseIf Not alreadysh.exists(oProduct.Products.Item(j).PartNumber) Then
    ShowallBodies2 oProduct.Products.Item(j).ReferenceProduct.Parent.Part, reverse
    alreadysh.Add oProduct.Products.Item(j).PartNumber, ""
    End If
On Error GoTo 0
Else
ShowallBodies oProduct.Products.Item(j), reverse
End If
oSel.Clear
Next
oProduct.ApplyWorkMode DEFAULT_MODE
End Sub
Sub ShowallBodies2(oPart, Optional reverse As Boolean = False)
Set oSel = oActDoc.Selection
oSel.Clear
oSel.Add oPart
oSel.VisProperties.SetShow 0
oSel.Clear
On Error Resume Next
CATIA.RefreshDisplay = False
Dim oBodies, oBody, k, showstate
 
                    Set oBodies = oPart.Bodies
                        For k = 1 To oBodies.Count
                            Set oBody = oBodies.Item(k)
                            oSel.Add oBody
                                If reverse = True Then
                                    oSel.VisProperties.GetShow showstate
                                    Select Case showstate
                                        Case 0
                                        oSel.VisProperties.SetShow 1
                                        Case 1
                                        oSel.VisProperties.SetShow 0
                                    End Select
                                Else
                                oSel.VisProperties.SetShow 0
                                End If
                            oSel.Clear
                        Next
CATIA.RefreshDisplay = True
End Sub
