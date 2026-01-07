Attribute VB_Name = "OTH_PrePn"
'Attribute VB_Name = "m30_PrePn"
'{GP:6}
'{Ep:CATMain}
'{Caption:零件号前缀}
'{ControlTipText:为所有零件号增加项目前缀}
'{BackColor:}

Private prj
Sub CATMain()

If Not KCL.CanExecute("ProductDocument") Then Exit Sub
    Set rootprd = CATIA.ActiveDocument.Product
If Not rootprd Is Nothing Then
 Dim imsg
          imsg = "请输入你的项目名称"
        prj = KCL.GetInput(imsg)
        If prj = "" Then
            Exit Sub
        End If
    Call rePn(rootprd)
Else
 Exit Sub
End If
End Sub

Sub rePn(oprd)
    pn = oprd.PartNumber
    purePN = KCL.straf1st(pn, "_")
    oprd.PartNumber = prj & "_" & purePN
    For Each Product In oprd.Products
        Call rePn(Product)
        Next
End Sub



Sub shot()
MsgBox "没编呢"
Exit Sub
 Dim iprd, rootprd, oprd, children
 Dim xlsht, rng, RC(0 To 1), oArry()
 Dim i, oRowNb
  RC(0) = 3: RC(1) = 3
    On Error Resume Next
    Set CATIA = GetObject(, "CATIA.Application") '获取catia程序
    Dim oDoc: Set oDoc = CATIA.ActiveDocument
    Set rootprd = CATIA.ActiveDocument.Product
         If Err.Number <> 0 Then
            MsgBox "请打开CATIA并打开你的产品，再运行本程序": Err.Clear
            Exit Sub
         End If
    On Error GoTo 0
    Set xlAPP = GetObject(, "Excel.Application") '获取excel程序
    Set xlsht = xlAPP.ActiveSheet: xlsht.Columns(2).NumberFormatLocal = "0.000"
Dim oWindow, oViewer
Dim file_type As String
Set oWindow = CATIA.ActiveWindow
oWindow.Layout = catWindowGeomOnly
Set oViewer = oWindow.ActiveViewer
oViewer.Reframe
'====修改背景颜色=====
Dim MyViewer, oColor(2)
Set MyViewer = CATIA.ActiveWindow.ActiveViewer
MyViewer.GetBackgroundColor oColor
MyViewer.PutBackgroundColor Array(1, 1, 1) ' Change background color to WHITE
'====修改背景颜色=====
file_type = "tiff"
Dim oname, CapturePath, oType
  CapturePath = CATIA.FileSelectionBox("输入文件名", file_type, CatFileSelectionModeSave)
  oname = CapturePath & "." & file_type
oType = catCaptureFormatTIFF 'catCaptureFormatBMP catCaptureFormatJPEG
MyViewer.CaptureToFile oType, oname ' MAIN SENTENCE!! STORE THE PICTURE IN ANY FORMAT
MyViewer.PutBackgroundColor oColor ' Change background original
MsgBox ("已经保存图片")
oWindow.Layout = catWindowSpecsAndGeom 'catWindowSpecsOnly ' catWindowGeomOnly
End Sub
Function shotme()
    Dim iprd, rootprd, oprd, children
    Dim xlsht, rng, RC(0 To 1), oArry()
    Dim i, oRowNb
     RC(0) = 3: RC(1) = 3
       On Error Resume Next
       Set CATIA = GetObject(, "CATIA.Application") '获取catia程序
       Dim oDoc: Set oDoc = CATIA.ActiveDocument
       Set rootprd = CATIA.ActiveDocument.Product
            If Err.Number <> 0 Then
               MsgBox "请打开CATIA并打开你的产品，再运行本程序": Err.Clear
               Exit Sub
            End If
    On Error GoTo 0
    Set xlAPP = GetObject(, "Excel.Application") '获取excel程序
    Set xlsht = xlAPP.ActiveSheet: xlsht.Columns(2).NumberFormatLocal = "0.000"
    Dim oWindow, oViewer
    Dim file_type As String
    Set oWindow = CATIA.ActiveWindow
    oWindow.Layout = catWindowGeomOnly
    Set oViewer = oWindow.ActiveViewer
    oViewer.Reframe
'====修改背景颜色=====
    Dim MyViewer, oColor(2)
    Set MyViewer = CATIA.ActiveWindow.ActiveViewer
    MyViewer.GetBackgroundColor oColor
    MyViewer.PutBackgroundColor Array(1, 1, 1) ' Change background color to WHITE
'====修改背景颜色=====
    file_type = "tiff"
    Dim oname, CapturePath, oType
    MyViewer.CaptureToClipboard
      CapturePath = CATIA.FileSelectionBox("输入文件名", file_type, CatFileSelectionModeSave)
      oname = CapturePath & "." & file_type
    oType = catCaptureFormatTIFF 'catCaptureFormatBMP catCaptureFormatJPEG
    MyViewer.CaptureToFile oType, oname ' MAIN SENTENCE!! STORE THE PICTURE IN ANY FORMAT
    MyViewer.PutBackgroundColor oColor ' Change background original
    MsgBox ("已经保存图片")
    oWindow.Layout = catWindowSpecsAndGeom 'catWindowSpecsOnly ' catWindowGeomOnly
End Function
