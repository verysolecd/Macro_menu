Attribute VB_Name = "OTH_shot"
'Attribute VB_Name = "m5_Cbom"
'{GP:6==}
'{Ep:shot}
'{Caption:��ͼ���ļ���}
'{ControlTipText:ѡ��Ҫ����ȡ���޸ĵĲ�Ʒ}
'{BackColor:16744703}

Sub shot()

MsgBox "û����"
Exit Sub

 Dim iprd, rootprd, oprd, children
 Dim xlsht, rng, RC(0 To 1), oArry()
 Dim i, oRowNb
  RC(0) = 3: RC(1) = 3
    On Error Resume Next
    Set CATIA = GetObject(, "CATIA.Application") '��ȡcatia����
    Dim oDoc: Set oDoc = CATIA.ActiveDocument
    Set rootprd = CATIA.ActiveDocument.Product
         If Err.Number <> 0 Then
            MsgBox "���CATIA������Ĳ�Ʒ�������б�����": Err.Clear
            Exit Sub
         End If
    On Error GoTo 0
    Set xlAPP = GetObject(, "Excel.Application") '��ȡexcel����
    Set xlsht = xlAPP.ActiveSheet: xlsht.Columns(2).NumberFormatLocal = "0.000"

Dim oWindow, oViewer
Dim file_type As String
Set oWindow = CATIA.ActiveWindow
oWindow.Layout = catWindowGeomOnly
Set oViewer = oWindow.ActiveViewer
oViewer.Reframe

'====�޸ı�����ɫ=====
Dim MyViewer, oColor(2)
Set MyViewer = CATIA.ActiveWindow.ActiveViewer
MyViewer.GetBackgroundColor oColor
MyViewer.PutBackgroundColor Array(1, 1, 1) ' Change background color to WHITE

'====�޸ı�����ɫ=====
file_type = "tiff"
Dim oName, CapturePath, oType
  CapturePath = CATIA.FileSelectionBox("�����ļ���", file_type, CatFileSelectionModeSave)
  oName = CapturePath & "." & file_type
oType = catCaptureFormatTIFF 'catCaptureFormatBMP catCaptureFormatJPEG
MyViewer.CaptureToFile oType, oName ' MAIN SENTENCE!! STORE THE PICTURE IN ANY FORMAT
MyViewer.PutBackgroundColor oColor ' Change background original
MsgBox ("�Ѿ�����ͼƬ")
oWindow.Layout = catWindowSpecsAndGeom 'catWindowSpecsOnly ' catWindowGeomOnly

End Sub
Function shotme()

    Dim iprd, rootprd, oprd, children
    Dim xlsht, rng, RC(0 To 1), oArry()
    Dim i, oRowNb
     RC(0) = 3: RC(1) = 3
       On Error Resume Next
       Set CATIA = GetObject(, "CATIA.Application") '��ȡcatia����
       Dim oDoc: Set oDoc = CATIA.ActiveDocument
       Set rootprd = CATIA.ActiveDocument.Product
            If Err.Number <> 0 Then
               MsgBox "���CATIA������Ĳ�Ʒ�������б�����": Err.Clear
               Exit Sub
            End If
    On Error GoTo 0
    Set xlAPP = GetObject(, "Excel.Application") '��ȡexcel����
    Set xlsht = xlAPP.ActiveSheet: xlsht.Columns(2).NumberFormatLocal = "0.000"

    Dim oWindow, oViewer
    Dim file_type As String
    Set oWindow = CATIA.ActiveWindow
    oWindow.Layout = catWindowGeomOnly
    Set oViewer = oWindow.ActiveViewer
    oViewer.Reframe

'====�޸ı�����ɫ=====
    Dim MyViewer, oColor(2)
    Set MyViewer = CATIA.ActiveWindow.ActiveViewer
    MyViewer.GetBackgroundColor oColor
    MyViewer.PutBackgroundColor Array(1, 1, 1) ' Change background color to WHITE

'====�޸ı�����ɫ=====
    file_type = "tiff"
    
    Dim oName, CapturePath, oType
    
    
    MyViewer.CaptureToClipboard
    
      CapturePath = CATIA.FileSelectionBox("�����ļ���", file_type, CatFileSelectionModeSave)
      oName = CapturePath & "." & file_type
      
    oType = catCaptureFormatTIFF 'catCaptureFormatBMP catCaptureFormatJPEG
    
    MyViewer.CaptureToFile oType, oName ' MAIN SENTENCE!! STORE THE PICTURE IN ANY FORMAT
    
    MyViewer.PutBackgroundColor oColor ' Change background original
    
    MsgBox ("�Ѿ�����ͼƬ")
    oWindow.Layout = catWindowSpecsAndGeom 'catWindowSpecsOnly ' catWindowGeomOnly



End Function





