Attribute VB_Name = "OTH_PrePn"
'Attribute VB_Name = "m30_PrePn"
'{GP:6}
'{Ep:Pnmgr}
'{Caption:零件号管理}
'{ControlTipText:零件号批量管理}
'{BackColor:}

'控件格式为 %UI <ControlType> <ControlName> <Caption/Text>
' %UEI Label lbL_jpzcs  键盘造车手出品
' %UI TextBox  txt_str 字符串
' %UI CheckBox chk_prefix  字符串增加为前缀
' %UI CheckBox  chk_suffix  字符串增加为后缀
' %UI CheckBox chk_delete  删除零件号内字符串
' %UI Button btnOK  确定
' %UI Button btncancel  取消


Private prj
Private Const mdlname As String = "OTH_PrePn"
Sub Pnmgr()
    If Not KCL.CanExecute("ProductDocument") Then Exit Sub
    Dim oprd:    Set oprd = CATIA.ActiveDocument.Product
    If oprd Is Nothing Then Exit Sub
    Dim oFrm: Set oFrm = KCL.newFrm(mdlname): oFrm.Show
    Select Case oFrm.BtnClicked
        Case "btnOK"
            istr = ""
           If oFrm.Res("txt_str") <> "" And Not KCL.ExistsKey(oFrm.Res("txt_str"), "字符") Then istr = oFrm.Res("txt_str")
           If istr = "" Then Exit Sub
                If oFrm.Res("chk_prefix") Then
                    Call c_pn_Prefix(oprd, istr)
                ElseIf oFrm.Res("chk_suffix") Then
                    Call c_pn_suffix(oprd, istr)
                ElseIf oFrm.Res("chk_delete") Then
                    Call del_pn_midx(oprd, istr)
                End If
        Case Else: Exit Sub
    End Select

End Sub

Sub c_pn_Prefix(oprd, istr)
        pn = oprd.partNumber
        purePN = KCL.straf1st(pn, "_")
        oprd.partNumber = istr & "_" & purePN
   If oprd.Products.count > 0 Then
    For Each Product In oprd.Products
        Call c_pn_Prefix(Product, istr)
    Next
    End If
End Sub

Sub c_pn_suffix(oprd, istr)
    pn = oprd.partNumber
    oprd.partNumber = pn & "_" & istr
   If oprd.Products.count > 0 Then
    For Each Product In oprd.Products
        Call c_pn_suffix(Product, istr)
    Next
    End If
End Sub
Function del_pn_midx(oprd, istr)
        pn = oprd.partNumber
        oprd.partNumber = VBA.Replace(pn, istr, "")
  If oprd.Products.count > 0 Then
    For Each Product In oprd.Products
        Call del_pn_midx(Product, istr)
    Next
   End If
End Function






Sub shot()
MsgBox "没编呢"
Exit Sub
 Dim iprd, rprd, oprd, children
 Dim xlsht, rng, RC(0 To 1), oArry()
 Dim i, oRowNb
  RC(0) = 3: RC(1) = 3
    On Error Resume Next
    Set CATIA = GetObject(, "CATIA.Application") '获取catia程序
    Dim odoc: Set odoc = CATIA.ActiveDocument
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
    Dim iprd, rprd, oprd, children
    Dim xlsht, rng, RC(0 To 1), oArry()
    Dim i, oRowNb
     RC(0) = 3: RC(1) = 3
       On Error Resume Next
       Set CATIA = GetObject(, "CATIA.Application") '获取catia程序
       Dim odoc: Set odoc = CATIA.ActiveDocument
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
