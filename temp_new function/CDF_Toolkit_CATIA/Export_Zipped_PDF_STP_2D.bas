Attribute VB_Name = "Export_Zipped_PDF_STP_2D"
Option Explicit

Sub CATMain()
IntCATIA
On Error Resume Next
If (TypeName(oActDoc) <> "DrawingDocument") Then
    MsgBox "此命令只能在工程制图模块中运行", vbInformation, "Information"
    Exit Sub
End If
Dim oDocStp1
Set oDocStp1 = DwgLinkedDoc(oActDoc)
Dim scpath
scpath = oCATVBA_Folder
Dim tempfolder
tempfolder = oCATVBA_Folder("Temp")

CDF_Tool.ExportStp oDocStp1, tempfolder
CDF_Tool.ExportPDF oActDoc, tempfolder

Dim params(1)
params(0) = tempfolder & "\" & Replace(oActDoc.Name, ".", "_") & ".pdf"
params(1) = oActDoc.Path & "\" & Replace(oActDoc.Name, ".", "_") & "_" & Year(Now) & Right("0" & Month(Now), 2) & Right("0" & Day(Now), 2) & ".zip"
If CreateObject("Scripting.FileSystemObject").FileExists(params(1)) Then
Kill params(1)
End If
CATIA.SystemService.ExecuteScript scpath, catScriptLibraryTypeDirectory, "Export_Zipped_PDF_STP.catvbs", "CATMain", params


Select Case TypeName(oDocStp1)
       Case "PartDocument"
       params(0) = tempfolder & "\" & Replace(oDocStp1.Name, ".CATPart", "_") & "CATPart.stp"
       Case "ProductDocument"
       params(0) = tempfolder & "\" & Replace(oDocStp1.Name, ".CATProduct", "_") & "CATProduct.stp"
End Select
CATIA.SystemService.ExecuteScript scpath, catScriptLibraryTypeDirectory, "Export_Zipped_PDF_STP.catvbs", "CATMain", params

End Sub
'CopyRight by 唐国庆Charles.Tang@2023-02-10
