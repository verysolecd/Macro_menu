Attribute VB_Name = "DRW_ExPDF"
'{GP:5}
'{EP:ExportPDF}
'{Caption:导出PDF}
'{ControlTipText: 一键导出PDF}
'{背景颜色: 12648447}

' %UI Label lbL_jpzcs  键盘造车手出品
' %UI CheckBox chk_path  是否导出到当前路径
' %UI CheckBox  chk_save  是否保存当前图纸
' %UI Button btnOK  确定
' %UI Button btncancel  取消

'------------------------------------------------
Option Explicit

Private Const mdlname As String = "DRW_ExPDF"
Sub ExportPDF()
If Not CanExecute("DrawingDocument") Then Exit Sub
'On Error Resume Next ' 临时开启错误处理
 Err.Number = 0
 Dim oEng: Set oEng = KCL.newEngine(mdlname): oEng.Show
 If LCase(oEng.ClickedButton) <> "btnok" Then Exit Sub
    Dim oDoc: Set oDoc = CATIA.ActiveDocument
    Dim opath As String: opath = ""
    If oEng.Results("chk_path") Then
        opath = IIf(oDoc.path = "", "", oDoc.path)
    Else:
        opath = KCL.selFdl()
    End If
    If opath = "" Then Exit Sub
        Dim DraftMgr: Set DraftMgr = CATIA.SettingControllers.item("DraftingOptions")
        Dim currSet: currSet = DraftMgr.GetAttr("DimDesignMode")
        DraftMgr.PutAttr "DimDesignMode", False
        DraftMgr.Commit
   If oEng.Results("chk_save") Then oDoc.Save
        Dim filePath(2) '0=路径，1=name，2=extname
        filePath(0) = opath
        filePath(1) = Replace(UCase(oDoc.Name), UCase(".CATDrawing"), "_") & KCL.timestamp("day") & "_"
        filePath(2) = "pdf"
        Dim pdfpath As String: pdfpath = KCL.JoinPathName(filePath)
        oDoc.ExportData pdfpath, "pdf"
        KCL.SmartOPenPath pdfpath
    DraftMgr.PutAttr "DimDesignMode", currSet
    DraftMgr.Commit
On Error GoTo 0
End Sub


