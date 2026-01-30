Attribute VB_Name = "A0TEST"
Sub GetActiveInWorkInfo()
    Dim odoc As Document
    Set odoc = CATIA.ActiveDocument
    Dim oPart As part
    Set oPart = GetCurrentPart(odoc)
    If oPart Is Nothing Then
        MsgBox "未能找到激活的零件。" & vbCrLf & _
               "如果你在总成环境下，请确保你在特征树上选中了该零件内的任意元素。", vbExclamation
        Exit Sub
    End If
    MsgBox "当前激活的零件是: " & oPart.Name
    Dim oInWorkObj As AnyObject
    Set oInWorkObj = oPart.InWorkObject
    If Not oInWorkObj Is Nothing Then
        MsgBox "当前定义的工作对象 (InWorkObject) 是: " & oInWorkObj.Name
    Else
        MsgBox "该零件没有激活的 InWorkObject。"
    End If
End Sub

' ---------------------------------------------------------
' 通用函数：尝试从当前文档或选择中获取 Part 对象
' ---------------------------------------------------------
Function GetCurrentPart(odoc As Document) As part
    On Error Resume Next

    If TypeName(odoc) = "PartDocument" Then
        Set GetCurrentPart = odoc.part
        Exit Function
    End If

    If TypeName(odoc) = "ProductDocument" Then
        Dim oSel As Selection
        Set oSel = odoc.Selection
        If oSel.count = 0 Then
             Set GetCurrentPart = Nothing
            Exit Function
        End If
        Dim i As Integer
        Dim tempObj As AnyObject
        Set tempObj = oSel.item(1).value
        Do While Not tempObj Is Nothing
            If TypeName(tempObj) = "Part" Then
                Set GetCurrentPart = tempObj
                Exit Function
            End If
               If TypeName(tempObj) = "Application" Or TypeName(tempObj) = "ProductDocument" Then
                Exit Do
            End If
            Set tempObj = tempObj.Parent
            Err.Clear
        Loop
    End If
    Set GetCurrentPart = Nothing
End Function




