Attribute VB_Name = "MDL_Bodyrename"
'{GP:4}
'{EP:bdyname}
'{Caption:实体重命名}
'{ControlTipText: 提示选择几实体后将实体按顺序重命名}
'{BackColor: }
'type definition
Option Explicit
Private Const mdlname As String = "MDL_Bodyrename"
Sub bdyname()
    If CATIA.Windows.count < 1 Then
        MsgBox "没有打开的窗口"
        Exit Sub
    End If
    
    On Error Resume Next
    Dim oDoc: Set oDoc = CATIA.ActiveDocument
    On Error GoTo 0
    
    Dim oprt
    Set oprt = KCL.get_workPartDoc.part
    If IsNothing(oprt) Then Exit Sub
    
    Dim osel: Set osel = CATIA.ActiveDocument.Selection
    If osel.count = 0 Then
        Set osel = KCL.Selectmulti("请选择BODY")
    End If
    
    Dim lst: Set lst = KCL.Initlst
    Dim itm, itp, i
    For i = 1 To osel.count
        Set itm = osel.item(i).Value
        Set itp = Nothing
        Set itp = KCL.GetParent_Of_T(itm, "Body")
        
        If Not itp Is Nothing Then
            lst.Add itp
        Else
            On Error Resume Next
                Dim itype: itype = TypeName(itm)
                Error.Clear
            On Error GoTo 0
        End If
        
        If LCase(itype) = LCase("Body") Then lst.Add itm
    Next i
    
    osel.Clear
    Set itm = Nothing
    
    Dim ct(0)
    ct(0) = 1
    For Each itm In lst
        If itm.InBooleanOperation = False Then
            itm.Name = "Body." & ct(0)
            ct(0) = ct(0) + 1
        End If
    Next
End Sub

