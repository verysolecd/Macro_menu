Attribute VB_Name = "Random_Color_3D"
#If VBA7 Then
Private Declare PtrSafe Function GetKeyState Lib "user32" (ByVal nVirtKey As Long) As Integer
'Private Declare PtrSafe Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer
#Else
Private Declare Function GetKeyState Lib "user32" (ByVal nVirtKey As Long) As Integer
'Private Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer
#End If
Dim ResetCol As Boolean


Sub CATMain()
       
IntCATIA
ResetCol = False
If (GetKeyState(vbKeyShift) And &H8000&) Then
    ResetCol = True
End If


'-----------设置设计模式-------------
CATIA.DisplayFileAlerts = False

Dim oProdxx
Set oProdxx = oActDoc.Product
oProdxx.ApplyWorkMode DESIGN_MODE '2
'-----------设置设计模式-------------

RndColor oActDoc
CATIA.DisplayFileAlerts = True
End Sub
Sub RndColor(oDocProd1)

If TypeName(oActDoc) = "ProductDocument" Then
Set oSel = oActDoc.Selection
RndCol (oActDoc.Product)
ElseIf TypeName(oActDoc) = "PartDocument" Then
Set oSel = oActDoc.Selection
RndCol2 oActDoc.Part
ElseIf TypeName(oActDoc) = "ProcessDocument" Then
Set oSel = oActDoc.Selection
RndCol oActDoc.PPRDocument.Products.Item(1)
Else
MsgBox "此命令不能在该工作台中运行!" & vbCrLf & "请在装配设计或零件设计工作台运行此命令！"
Exit Sub
End If
End Sub

Sub RndCol(oProduct)
oProduct.ApplyWorkMode DESIGN_MODE
oSel.Clear
oSel.Add oProduct
oSel.VisProperties.SetRealColor 255, 255, 255, 0
oSel.Clear

Dim j
For j = 1 To oProduct.Products.Count

If oProduct.Products.Item(j).Products.Count = 0 Then
On Error Resume Next
    oProduct.Products.Item(j).ApplyWorkMode DESIGN_MODE
    oSel.Add oProduct.Products.Item(j)
    oSel.VisProperties.SetRealColor 255, 255, 255, 0
    RndCol2 (oProduct.Products.Item(j).ReferenceProduct.Parent.Part)
On Error GoTo 0
Else
RndCol (oProduct.Products.Item(j))
End If

oSel.Clear
Next
oProduct.ApplyWorkMode DEFAULT_MODE
End Sub

Sub RndCol2(oPart)
Dim R, G, b
oSel.Clear
oSel.Add oPart
oSel.VisProperties.SetRealColor 255, 255, 255, 0

'*****************************
Dim oBodies, oBody, k
 
                    Set oBodies = oPart.Bodies
                        For k = 1 To oBodies.Count
                            Set oBody = oBodies.Item(k)
                            oSel.Add oBody
                            If (ResetCol = False) And (oBody.InBooleanOperation = False) Then
                                Randomize
                                R = CLng(255 * Rnd)
                                Randomize
                                G = CLng(255 * Rnd)
                                Randomize
                                b = CLng(255 * Rnd)
                                oSel.VisProperties.SetRealColor R, G, b, 1
                                
                            ElseIf ResetCol = False Then
                                oSel.VisProperties.SetRealColor R, G, b, 1
                            Else
                                oSel.VisProperties.SetRealColor 210, 210, 255, 1
                                oSel.VisProperties.SetRealOpacity 255, 1
                            End If
                            oSel.Clear

                            Dim shapes1 'As Shapes
                            Dim m As Integer
                            Set shapes1 = oBody.Shapes
                            For m = 1 To shapes1.Count
                            'MsgBox TypeName(shapes1.Item(m))
                             If TypeName(shapes1.Item(m)) = "Solid" Then
                                oSel.Add shapes1.Item(m)
                                If (ResetCol = False) And (oBody.InBooleanOperation = False) Then
                                    oSel.VisProperties.SetRealColor R, G, b, 1
                                ElseIf ResetCol = False Then
                                    oSel.VisProperties.SetRealColor R, G, b, 1
                                Else
                                    oSel.VisProperties.SetRealColor 210, 210, 255, 1
                                    oSel.VisProperties.SetRealOpacity 255, 1
                                End If
                                oSel.Clear
                            End If
                            Next
                        Next


End Sub
