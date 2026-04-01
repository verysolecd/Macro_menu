Attribute VB_Name = "ZZ_BCK"
Private Declare PtrSafe Function OpenClipboard Lib "user32" (ByVal hwnd As LongPtr) As Long
Private Declare PtrSafe Function EmptyClipboard Lib "user32" () As Long
Private Declare PtrSafe Function CloseClipboard Lib "user32" () As Long
Private Declare PtrSafe Function SetClipboardData Lib "user32" (ByVal wFormat As Long, ByVal hMem As LongPtr) As LongPtr
Private Const mdlname As String = "ZZ_BCK"
Private Sub remove_usrP()
    Set oPrd = CATIA.ActiveDocument.Product
    rm oPrd
End Sub
Private Sub rm(oPrd)
    On Error Resume Next
     Set refPrd = oPrd.ReferenceProduct
     Set oprt = refPrd.Parent.part
    Set colls = refPrd.Publications
    colls.Remove ("Location")
    colls.Remove ("iMass")
    colls.Remove ("iDensity")
    colls.Remove ("iThickness")
    colls.Remove ("iMaterial")
     Set colls = refPrd.Parent.part.Parameters.RootParameterSet.ParameterSets
        Set cm = colls.GetItem("cm")
        Set osel = CATIA.ActiveDocument.Selection
        osel.Clear: osel.Add cm: osel.Delete
     Set colls = refPrd.Parent.part.Relations
     colls.Remove ("CalM")
     colls.Remove ("CMAS")
     colls.Remove ("CTK")
     Set colls = refPrd.UserRefProperties
     colls.Remove ("iMass")
     colls.Remove ("iMaterial")
     colls.Remove ("iThickness")
    If oPrd.Products.count > 0 Then
        For i = 1 To oPrd.Products.count
          rm (oPrd.Products.item(i))
        Next
    End If
On Error GoTo 0
End Sub


''==遍历递归=============================
Private Sub recurAyo(ayo)
    Dim colls: Set itm = ayo.Products
    For Each itm In colls
        Call recurFunc(itm)
    Next

    If ayo.Products.count > 0 Then
            For Each ctm In ayo.Products
                Call recurAyo(ctm)
             Next
    End If
End Sub

''==图纸页面=============================

Private Sub main()
CATIA.RefreshDisplay = False
    Set shts = CATIA.ActiveDocument.sheets
      Set osht = Nothing
    Set lst = InitDic()
j = 1
       For i = 1 To shts.count
           Set osht = shts.item(i)
               If osht.IsDetail = False Then

                 lst.Add j, osht
        j = j + 1
               End If
       Next
    Set osht = Nothing
    For i = 1 To lst.count
       Set osht = lst(i)
       If osht.IsDetail = False Then
            osht.Activate
                    oo = StrAF(osht.name, " ")
        If i > 9 Then
            osht.name = "SH" & i & oo
        Else
             osht.name = "SH0" & i & oo
        End If
            Set oView = osht.Views.item("Background View")
            Set ots = oView.Texts
            Set oDict = InitDic()
            For Each itm In ots
               Set oDict(itm.name) = itm
            Next
           Set Pg1 = oDict("gongxxzhang")
            Pg1.text = "共" & shts.count - 1 & "页"
            Set Pg2 = oDict("dixxzhang")
            Pg2.text = "第" & i & "页"
            oView.SaveEdition
        End If
    Next
     CATIA.RefreshDisplay = True
     Set oView = osht.Views.item(1)
      osht.Activate
End Sub
Private Function StrAF(istr, iext)
Dim idx
idx = InStr(istr, iext)
If idx > 0 Then
        StrAF = Mid(istr, idx)
    Else
        StrAF = istr
    End If
End Function

Private Function InitDic()
    Dim dic As Object
    Set dic = CreateObject("Scripting.Dictionary")
    dic.compareMode = compareMode
    Set InitDic = dic
End Function


Sub shot()
MsgBox "没编呢"
Exit Sub
 Dim iprd, rprd, oPrd, children
 Dim xlsht, rng, RC(0 To 1), oArry()
 Dim i, oRowNb
  RC(0) = 3: RC(1) = 3
    On Error Resume Next
    Set CATIA = GetObject(, "CATIA.Application") '获取catia程序
    Dim oDoc: Set oDoc = CATIA.ActiveDocument
    Set rprd = CATIA.ActiveDocument.Product
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
    Dim iprd, rprd, oPrd, children
    Dim xlsht, rng, RC(0 To 1), oArry()
    Dim i, oRowNb
     RC(0) = 3: RC(1) = 3
       On Error Resume Next
       Set CATIA = GetObject(, "CATIA.Application") '获取catia程序
       Dim oDoc: Set oDoc = CATIA.ActiveDocument
       Set rprd = CATIA.ActiveDocument.Product
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




