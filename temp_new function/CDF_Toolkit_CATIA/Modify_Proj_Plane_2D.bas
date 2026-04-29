Attribute VB_Name = "Modify_Proj_Plane_2D"
Sub CATMain()
On Error Resume Next
IntCATIA
'MsgBox CATIA.Windows.Item(1).Caption
Select Case TypeName(oActDoc)
    Case "DrawingDocument"
    Case Else
        MsgBox "只能在工程制图工作台操作这个命令"
    Exit Sub
End Select

Dim V5Lang As String
V5Lang = GetV5Lang
Debug.Print "V5界面语言是 " & V5Lang
Dim oMyWindow As Object
Dim o3DWindow As Object
Dim sMyCaption As String
Dim s3DCaption As String
Dim FlagFind As Boolean
FlagFind = False
sMyCaption = CATIA.ActiveWindow.Caption
Dim oMyView As Object
Err.Clear
           On Error Resume Next
                Set oMyView = Sel("DrawingView")
                If Err.Number <> 0 Then
                Exit Sub
                End If
                   Select Case oMyView.ViewType
                            Case catViewFront
                            Case Else
                            MsgBox "只能选择主视图", vbInformation
                            Exit Sub
                   End Select
oMyView.Activate
'MsgBox TypeName(oMyView.GenerativeBehavior.Document.Parent)
Select Case TypeName(oMyView.GenerativeBehavior.Document.Parent) '多实体模式是Bodies
    Case "Bodies"
        s3DCaption = oMyView.GenerativeBehavior.Document.Parent.Parent.Parent.Name  '如果是多实体模式需要多两个Parent
        '查询数模窗口是否打开
        Set oMyWindow = CATIA.ActiveWindow
        For Each o3DWindow In CATIA.Windows
        If o3DWindow.Caption = s3DCaption Then
        FlagFind = True
        Exit For
        End If
        Next
        
        If FlagFind = False Then
        CATIA.Documents.Open oMyView.GenerativeBehavior.Document.Parent.Parent.Parent.FullName
        Set o3DWindow = CATIA.ActiveWindow
        FlagFind = True
        End If
        oMyWindow.Activate
        oMyView.Activate
        Set oSel = oActDoc.Selection
        oSel.Clear
        oSel.Add oMyView
    Case Else
        s3DCaption = oMyView.GenerativeBehavior.Document.Parent.Name
        '查询数模窗口是否打开
        Set oMyWindow = CATIA.ActiveWindow
        For Each o3DWindow In CATIA.Windows
        If o3DWindow.Caption = s3DCaption Then
        FlagFind = True
        Exit For
        End If
        Next
        
        If FlagFind = False Then
        CATIA.Documents.Open oMyView.GenerativeBehavior.Document.Parent.FullName
        Set o3DWindow = CATIA.ActiveWindow
        FlagFind = True
        End If
        oMyWindow.Activate
        oMyView.Activate
        Set oSel = oActDoc.Selection
        oSel.Clear
        oSel.Add oMyView
End Select


If V5Lang = "English" Then
CATIA.StartCommand "Modify Projection Plane"
Debug.Print "执行 Modify Projection Plane"
ElseIf V5Lang = "Simplified Chinese" Then
CATIA.StartCommand "修改投影平面"
Debug.Print "执行 修改投影平面"
End If
o3DWindow.Activate
End Sub
Function GetV5Lang() As String
Dim s As Object
Set s = CATIA.ActiveDocument.Selection
s.Clear
Dim stu As String
stu = CATIA.StatusBar
stu = LTrim(stu)
stu = Mid(stu, 1, 1) '第一个字符
Select Case Asc(UCase(stu))
Case 65 To 90           'A-Z
GetV5Lang = "English"
Case Else
GetV5Lang = "Simplified Chinese"
End Select
End Function
Function Sel(objType As String, Optional ObjType2 As String, Optional ObjType3 As String)
Dim s2, InputObjectType(), Status, NotSel, obj, Picked
Set s2 = CATIA.ActiveDocument.Selection
s2.Clear

If ObjType2 = "" Then
        ReDim InputObjectType(0)
        InputObjectType(0) = objType
        ElseIf ObjType3 = "" Then
                ReDim InputObjectType(1)
                InputObjectType(0) = objType
                InputObjectType(1) = ObjType2
                Else
                    ReDim InputObjectType(2)
                    InputObjectType(0) = objType
                    InputObjectType(1) = ObjType2
                    InputObjectType(2) = ObjType3
End If

Dim indication, n
indication = ""
For n = 0 To UBound(InputObjectType)
    indication = indication & "," & InputObjectType(n)
Next

Picked = False
NotSel = True
    Do While NotSel
        Status = s2.SelectElement2(InputObjectType, "Select the " & indication, False)
        If (Status = "Cancel") Then
            Set Sel = Nothing 'modified 20200616
            MsgBox "已退出选择", vbSystemModal
            Exit Function
        ElseIf (Status = "Redo" And Not Picked) Then
               ElseIf (Status = "Undo") Then
                        Exit Function
                    ElseIf (Status <> "Redo") Then Set obj = s2.Item(1).Value
                              Picked = True
                              NotSel = False
        End If
    Loop
    s2.Clear
    Set Sel = obj
End Function
