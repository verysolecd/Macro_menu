Attribute VB_Name = "Open_Drawing_Model_3D"
Sub CATMain()
IntCATIA
On Error Resume Next
If oActDoc Is Nothing Then
Exit Sub
End If
If TypeName(oActDoc) <> "ProductDocument" And TypeName(oActDoc) <> "PartDocument" And TypeName(oActDoc) <> "DrawingDocument" Then
MsgBox "当前活动文档不是.CATProduct或.CATPart或.CATDrawing!"
Exit Sub
End If

Dim shortName As String
Debug.Print TypeName(oActDoc)
Select Case TypeName(oActDoc)
Case "ProductDocument"
     shortName = Replace(oActDoc.Name, ".CATProduct", "")
Case "PartDocument"
     shortName = Replace(oActDoc.Name, ".CATPart", "")
Case "DrawingDocument"
     CATIA.Documents.Open DwgLinkedDoc(oActDoc).FullName
     Exit Sub
Case Else
End Select

Dim fso1, file1, n1
n1 = 0
Dim arr() As String
Set fso1 = CreateObject("Scripting.FileSystemObject")
For Each file1 In fso1.GetFolder(oActDoc.Path).Files
    If (InStr(file1.Name, ".CATDrawing") > 0) And (InStr(file1.Name, shortName) > 0) Then
    ReDim Preserve arr(n1)
    arr(n1) = oActDoc.Path & "\" & file1.Name
    n1 = n1 + 1
    End If
Next
If n1 = 1 Then
CATIA.Documents.Open arr(n1 - 1)
ElseIf n1 > 1 Then
    Dim i As Integer
    Dim drwlist As String
    Dim shortDrwName As String
    For i = 0 To n1 - 1
    shortDrwName = Right(arr(i), Len(arr(i)) - InStrRev(arr(i), "\"))
    drwlist = drwlist & vbCrLf & shortDrwName
    Next
    If (MsgBox("当前文件夹下找到疑似工程图" & UBound(arr) + 1 & "个，需要手动打开工程图吗？" & vbCrLf & drwlist, vbQuestion + vbYesNo, "臭豆腐工具箱CATIA版") = vbYes) Then
    Open_Current_Folder.CATMain
    End If
ElseIf n1 < 1 Then
    If (MsgBox("当前文件夹下找不到对应工程图！需要手动打开文件夹吗？", vbQuestion + vbYesNo, "臭豆腐工具箱CATIA版") = vbYes) Then
    Open_Current_Folder.CATMain
    End If
End If
Set fso1 = Nothing
Set file1 = Nothing
End Sub
