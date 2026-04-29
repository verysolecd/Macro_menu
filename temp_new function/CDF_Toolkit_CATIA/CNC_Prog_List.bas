Attribute VB_Name = "CNC_Prog_List"
Option Explicit
Const intRowStart = 10   'contents start row
Const intColStart = 2  'contents start column
Dim oMyPPR As Document
Dim oDicF As Object  'Dictionary for contents
Dim ItemNo As Integer




Sub CATMain()
'
IntCATIA
If TypeName(oActDoc) <> "ProcessDocument" Then
    MsgBox "此命令只能在加工模块中运行！"
    Exit Sub
Else
    Set oMyPPR = oActDoc.PPRDocument
End If

Set oDicF = CreateObject("Scripting.Dictionary")

IntExcel
On Error Resume Next
Dim sExlTp As String
sExlTp = oCATVBA_Folder("Template").Path & "\" & "MProg.xls"
'Debug.Print sExlTp.Name
Set objExcel = objExcel.workbooks.Add(Template:=sExlTp)
    If Err.Number <> 0 Then
        MsgBox "无法取得Excel属性模板，程序退出"
        Exit Sub
    End If
'On Error GoTo 0
Err.Clear
CATIA.DisplayFileAlerts = False
objExcel.Parent.DisplayAlerts = False
'objExcel.Parent.ScreenUpdating = False
objExcel.Parent.ScreenUpdating = True
objExcel.Parent.Visible = True
Dim oMyPrcs As Object
Dim i As Integer
Set oMyPrcs = oMyPPR.processes      'ProcessList


' Capture Image
Dim FrontViewJpg As String
Dim TopViewJpg As String
FrontViewJpg = oCATVBA_Folder("Temp").Path & "\" & Replace(oActDoc.Name, ".", "_") & "_" & "FrontView.jpg"
TopViewJpg = oCATVBA_Folder("Temp").Path & "\" & Replace(oActDoc.Name, ".", "_") & "_" & "TopView.jpg"
CapImage FrontViewJpg, TopViewJpg

'MsgBox oMyPrcs.Count

For i = 1 To oMyPrcs.Count
    ProcessFinInfo oMyPrcs.Item(i)
    Dim ArrKeys
    ArrKeys = oDicF.keys
'    Set objExcel = objExcel.Sheets.Item("MP")  '工作表
     objExcel.Sheets.Item("MP").Visible = False
     objExcel.Sheets.Item("MP").cells.Copy
     Set objExcel = objExcel.Sheets.Add
     objExcel.Name = oMyPrcs.Item(i).Name
     objExcel.Paste
     
     
        Dim k As Integer
 
        For k = 0 To UBound(ArrKeys)

        objExcel.Range(objExcel.cells(intRowStart + k, intColStart), objExcel.cells(intRowStart + k, intColStart + UBound(oDicF.Item(ArrKeys(k))))).Value = oDicF.Item(ArrKeys(k))
        
        Next
'************************Format Excel*******************************
Dim Rng
Set Rng = objExcel.Range(objExcel.cells(intRowStart, intColStart), objExcel.cells(intRowStart + UBound(ArrKeys), intColStart + UBound(oDicF.Item(ArrKeys(1)))))
With Rng.Borders(7) '(xlEdgeLeft)
        .LineStyle = 1 ' xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = 2 ' xlThin
End With
With Rng.Borders(10) '(xlEdgeRight)
        .LineStyle = 1 ' xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = 2 ' xlThin
End With
With Rng.Borders(8) '(xlEdgeTop)
        .LineStyle = 1 ' xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = 2 ' xlThin
End With
With Rng.Borders(9) '(xlEdgeBottom)
        .LineStyle = 1 ' xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = 2 ' xlThin
End With
With Rng.Borders(12) '(xlInsideHorizonta)
        .LineStyle = 1 ' xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = 2 ' xlThin
End With
With Rng.Borders(11)  '(xlInsideHorizonta)
        .LineStyle = 1 ' xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = 2 ' xlThin
End With
'合并描述列
Set Rng = objExcel.Range(objExcel.cells(intRowStart, 5), objExcel.cells(intRowStart + UBound(ArrKeys), 7))
Rng.Merge True

Set Rng = Nothing
'************************Format Excel*******************************
On Error Resume Next
    objExcel.cells(11, 16).Select
    Dim Frontjpg, Topjpg
    Set Frontjpg = objExcel.Pictures.Insert(FrontViewJpg)
    Frontjpg.ShapeRange.LockAspectRatio = 1
    Frontjpg.ShapeRange.Height = 228
    objExcel.cells(22, 16).Select
    Set Topjpg = objExcel.Pictures.Insert(TopViewJpg)
    Topjpg.ShapeRange.LockAspectRatio = 1
    Topjpg.ShapeRange.Height = 228
    'Table Head
    objExcel.cells(2, 10).Value = oActDoc.FullName
    objExcel.cells(6, 3).Value = oMyPPR.Products.Item(1).PartNumber
    objExcel.cells(4, 3).Value = oMyPPR.Products.Item(1).Definition
    objExcel.cells(10, 16).Value = "输出项和表格格式定制，关注微信公众号【UG遇上CATIA】"
    oDicF.RemoveAll
On Error GoTo 0
Next


CATIA.DisplayFileAlerts = True
objExcel.Parent.Parent.DisplayAlerts = True
objExcel.Parent.Parent.ScreenUpdating = True
objExcel.Parent.Parent.Visible = True
Dim sFName
sFName = oCATVBA_Folder("NC_Files").Path & "\" & Replace(oActDoc.Name, ".", "_") & ".xlsm"
On Error Resume Next
objExcel.Parent.SaveAs sFName, 52
If Err.Number <> 0 Then
MsgBox "无法保存，是否已经打开同名Excel?"
End If

Set objExcel = Nothing
Set Frontjpg = Nothing
Set Topjpg = Nothing
Set oMyPrcs = Nothing
Set oDicF = Nothing

End Sub

Sub ProcessFinInfo(MyPs As Object) ' process activity
    Dim ii As Integer
    Dim childs As Object
    Dim child As Object
    On Error Resume Next
    Set childs = MyPs.ChildrenActivities
    If childs.Count <= 0 Then Exit Sub

  For ii = 1 To childs.Count
    'MsgBox childs.Item(ii).Name & ", TypeName is " & TypeName(childs.Item(ii))
    Select Case TypeName(childs.Item(ii))
        Case "ManufacturingSetup"                        'Part Operation
              If childs.Item(ii).Active Then
              PoToolList (childs.Item(ii))
             'MsgBox childs.Item(ii).Name & ", TypeName is " & TypeName(childs.Item(ii))
             End If
        Case "PPRActivity"
             If childs.Item(ii).Active Then
             Call ProcessFinInfo(childs.Item(ii))
             End If
    End Select
  Next '**ii
    On Error GoTo 0
End Sub

Sub PoToolList(po) 'po As ManufacturingSetup
    If TypeName(po) <> "ManufacturingSetup" Then
        Exit Sub
    End If
    Dim oMyPrgs As Object
    Dim iii As Integer
    Set oMyPrgs = po.Programs
    For iii = 1 To oMyPrgs.Count
    'MsgBox po.Name & "有" & oMyPrgs.Count & "个Programs,现在是第" & iii & "个"
         Dim oMyPrg As Object
         Set oMyPrg = oMyPrgs.GetElement(iii)
             'MsgBox oMyPrg.Name & ",TypeName is " & TypeName(oMyPrg)
             If oMyPrg.Active Then
             ToolInfo oMyPrg
             End If
    Next 'iii

End Sub
Sub ToolInfo(mp As Object) 'mp is program

    If TypeName(mp) <> "ManufacturingProgram" Then
        Exit Sub
    End If
On Error Resume Next
Dim Acts As Object
Dim iiii As Integer
Set Acts = mp.Activities

  For iiii = 1 To Acts.Count
   Dim Act As Object
   Set Act = Acts.GetElement(iiii)
'   MsgBox Act.Type



   If Act.Active Then
       If (Act.Type <> "ToolChange" And Act.Type <> "ToolChangeLathe" And Act.Type <> "TableHeadRotation" And Act.Type <> "CoordinateSystem" And Act.Type <> "PPInstruction" And _
       Act.Type <> "MfgTracutOperation" And Act.Type <> "MfgTracutEnd") Then
                ItemNo = ItemNo + 1
                Dim aTool As Object
                Set aTool = Act.Tool
                'MsgBox aTool.Number ' aTool.Name, aTool.ToolNumber,aTool.ToolType
                ' This VBA Macro Developed by Charles.Tang
                ' WeChat Chtang80,CopyRight reserved
                Dim ToolLength, ToolCornerRadius
                ToolLength = Val(aTool.GetAttribute("MFG_LENGTH").ValueAsString)
                ToolCornerRadius = "R" & Val(aTool.GetAttribute("MFG_CORNER_RAD").ValueAsString)
                Dim DiameterAttribut As Object
                Dim DiameterParameterName As String
                Dim ToolDiameter
                If (aTool.ToolType = "MfgAPTTool") Then
                        DiameterParameterName = "MFG_APT_DIAMETER"
                    Else
                        DiameterParameterName = "MFG_NOMINAL_DIAM"
                End If
                Err.Clear
                Set DiameterAttribut = aTool.GetAttribute(DiameterParameterName)
                
                If (Err.Number = 0) Then
                    ToolDiameter = "D" & Val(DiameterAttribut.ValueAsString)
                Else
                    ToolDiameter = "??"
                End If
                Dim ArrToolInfo()  '一维数组
                ReDim ArrToolInfo(12)
                'ArrToolInfo = Array("程序名", "直径/R角", "装刀长", "加工描述", "", "", "", "", "侧余量", "底余量", "Z深", "理论时间（秒）", "备注")
                'ArrToolInfo = Array("", "", "", "", "", "", "", "", "", "", "", "", "")
                ArrToolInfo(0) = mp.Name
                ArrToolInfo(1) = ToolDiameter & "/" & ToolCornerRadius
                ArrToolInfo(2) = ToolLength
                ArrToolInfo(3) = Act.Name
                Err.Clear
                    ArrToolInfo(10) = Act.ToolAssembly.GetAttribute("MFG_NAME").Value
                    If Err.Number <> 0 Then
                    ArrToolInfo(10) = ""
                    Err.Clear
                    End If
                ArrToolInfo(11) = Round(Act.TotalTime, 2)
                If Not oDicF.exists(ItemNo) Then
                oDicF.Add ItemNo, ArrToolInfo
                End If
                'MsgBox TypeName(oDicF.Item(ItemNo)) & vbCrLf & ArrToolInfo(0) & " | " & ArrToolInfo(1) & " | " & ArrToolInfo(2) & " | " & ArrToolInfo(3) & " | " & ArrToolInfo(11) & "秒" & vbCrLf & _
                       aTool.ToolNumber & "||" & aTool.Name & " || " & ToolLength & " || " & ToolCornerRadius & vbCrLf & ToolDiameter
        End If
   End If
  Next 'iiii
End Sub


Sub CapImage(sFullNameFront As String, sFullNameTop As String)  'sFullName must end with ".jpg"

    Dim specsAndGeomWindow1
    Set specsAndGeomWindow1 = CATIA.ActiveWindow
    Dim W, h
    W = specsAndGeomWindow1.Width
    h = specsAndGeomWindow1.Height
    specsAndGeomWindow1.Width = 400
    specsAndGeomWindow1.Height = 300

    Dim viewpoint3D1
    Set viewpoint3D1 = specsAndGeomWindow1.ActiveViewer.Viewpoint3D

    Dim OldSight(2)
    Dim OldUp(2)
    viewpoint3D1.GetSightDirection OldSight
    viewpoint3D1.GetUpDirection OldUp
    Dim myOldLayout, myViewer
    myOldLayout = specsAndGeomWindow1.Layout
    specsAndGeomWindow1.Layout = 1
    Set myViewer = specsAndGeomWindow1.ActiveViewer

    Dim BGcolor(3)
    myViewer.GetBackgroundColor BGcolor
    myViewer.PutBackgroundColor Array(1, 1, 1)


    'FrontView
        viewpoint3D1.PutSightDirection Array(0, 0, -1)
        viewpoint3D1.PutUpDirection Array(-1, 0, 0)
        myViewer.Reframe
    myViewer.CaptureToFile catCaptureFormatJPEG, sFullNameFront

     'TopView
        viewpoint3D1.PutSightDirection Array(-1, 0, 0)
        viewpoint3D1.PutUpDirection Array(0, 0, 1)
        myViewer.Reframe
     myViewer.CaptureToFile catCaptureFormatJPEG, sFullNameTop


    myViewer.PutBackgroundColor BGcolor

    specsAndGeomWindow1.Layout = myOldLayout

    viewpoint3D1.PutSightDirection OldSight
    viewpoint3D1.PutUpDirection OldUp

    specsAndGeomWindow1.Width = W
    specsAndGeomWindow1.Height = h
' This VBA Macro Developed by Charles.Tang
' WeChat Chtang80,CopyRight reserved
     myViewer.Reframe
     Set myViewer = Nothing
     Set viewpoint3D1 = Nothing
     Set specsAndGeomWindow1 = Nothing
End Sub
