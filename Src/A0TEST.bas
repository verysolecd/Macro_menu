Attribute VB_Name = "A0TEST"
Sub test()
    Set mprt = KCL.get_inwork_part
    Set mHSF = mprt.HybridShapeFactory
    Dim osel
    Set osel = CATIA.ActiveDocument.Selection  ' mprt.Parent.Selection '=
    Dim lst: Set lst = KCL.Initlst
    Dim itm

    Set itm = osel.item(1).value
    Set bb = Nothing
    Set bb = KCL.GetParent_Of_T(itm, "Body")
    If Not bb Is Nothing Then
    lst.Add bb
    Else
    On Error Resume Next
    Dim itype:  itype = mHSF.GetGeometricalFeatureType(itm)
    Error.Clear
    On Error GoTo 0
    End If
    If LCase(itype) = 7 Then lst.Add itm

MsgBox "d"
End Sub
Sub test_Debug()
    Dim osel As Selection
    Set osel = CATIA.ActiveDocument.Selection
      Set mprt = KCL.get_inwork_part
    Set mHSF = mprt.HybridShapeFactory
    If osel.count = 0 Then MsgBox "未选中任何对象": Exit Sub
    
    Dim i As Integer
    Dim itm As Object
    Dim selElem As SelectedElement
    
    For i = 1 To osel.count
        Set selElem = osel.item(i)
        Set itm = selElem.value

            Set bb = Nothing
            Set bb = KCL.GetParent_Of_T(itm, "Body")
         If Not bb Is Nothing Then
            lst.Add bb
         Else
            On Error Resume Next
                Dim itype:  itype = mHSF.GetGeometricalFeatureType(itm)
                Error.Clear
            On Error GoTo 0
         End If
        If itype = 7 Then lst.Add itm
        
        MsgBox "第 " & i & " 个选中项详情:" & vbCrLf & _
               "TypeName(.Value): " & TypeName(itm) & vbCrLf & _
               "Name(.Value): " & itm.Name & vbCrLf & _
               "Type(.Type): " & selElem.Type & vbCrLf & _
               "LeafProduct: " & selElem.LeafProduct.Name & _
                "lei" & itype
               
               
        ' 如果 TypeName 是 Part，说明确实选中了 Part 对象
        ' 此时请检查 CATIA 界面下方的 "User Selection Filter" 是否开启了特定过滤
    Next i
End Sub

