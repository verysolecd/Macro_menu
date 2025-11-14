VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} GB_Frame 
   Caption         =   "新建 |  输出图纸"
   ClientHeight    =   5985
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   2310
   OleObjectBlob   =   "GB_Frame.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "GB_Frame"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
#If VBA7 Then
Private Declare PtrSafe Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As LongPtr
Private Declare PtrSafe Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As LongPtr
#Else
Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Private Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long
#End If








Private Sub cmdNewfrmTemplate_Click()
New_From_Template.CATMain
End Sub



Private Sub GB_A0_Click()

Dim sTemplate
sTemplate = oCATVBA_Folder("Template") & "\A0.CATDrawing"
InsFrame sTemplate
Unload GB_Frame
End Sub

Private Sub GB_A1_Click()

Dim sTemplate
sTemplate = oCATVBA_Folder("Template") & "\A1.CATDrawing"
InsFrame (sTemplate)
Unload GB_Frame

End Sub

Private Sub GB_A2_Click()

Dim sTemplate
sTemplate = oCATVBA_Folder("Template") & "\A2.CATDrawing"
InsFrame (sTemplate)
Unload GB_Frame

End Sub

Private Sub GB_A3_Click()

Dim sTemplate
sTemplate = oCATVBA_Folder("Template") & "\A3.CATDrawing"
InsFrame (sTemplate)
Unload GB_Frame

End Sub
Private Sub GB_A4_Click()

Dim sTemplate
sTemplate = oCATVBA_Folder("Template") & "\A4.CATDrawing"
InsFrame (sTemplate)
Unload GB_Frame

End Sub

Private Sub GB_REFRESH_Click()

RefreshTitleBlock oActDoc
Unload GB_Frame
End Sub



'************************************
'************************************


Sub InsFrame(sTemplateFile)

On Error Resume Next
Dim sPartName, sPartNumber, sMass, sVolume, sMaterial
sPartName = ""
sPartNumber = ""
'sMaterial = "材料:" & vbCrLf ' & "" 'v021s修改
' This VBA Macro Developed by Charles.Tang
' WeChat Chtang80,CopyRight reserved
sMass = ""
sVolume = ""
sPartName = sPartName & oActDoc.Product.Definition 'v021s修改
sPartNumber = sPartNumber & oActDoc.Product.PartNumber
'sMaterial = sMaterial & oActDoc.Product.Nomenclature  'v021s增加
sMaterial = getMaterial(oActDoc.Product)
sMass = CStr(Format(Round(oActDoc.Product.Analyze.Mass, 3), "0.000")) & "kg"
sVolume = CStr(oActDoc.Product.Analyze.Volume)

Dim sDrwFileName As String
sDrwFileName = ""
sDrwFileName = oActDoc.FullName
Select Case TypeName(oActDoc)
    Case "PartDocument"
          sDrwFileName = Left(sDrwFileName, Len(sDrwFileName) - 8) & ".CATDrawing" '移除.CATPart
    Case "ProductDocument"
          sDrwFileName = Left(sDrwFileName, Len(sDrwFileName) - 11) & ".CATDrawing" '移除.CATProduct
    Case Else
End Select
If CreateObject("Scripting.FileSystemObject").FileExists(sDrwFileName) Then
   sDrwFileName = Left(sDrwFileName, Len(sDrwFileName) - 11) & "_" & CStr(Year(Now)) & CStr(Month(Now)) & CStr(Day(Now)) & ".CATDrawing"
End If

If CreateObject("Scripting.FileSystemObject").FileExists(sDrwFileName) Then
Randomize
   sDrwFileName = Left(sDrwFileName, Len(sDrwFileName) - 11) & "_" & Int(100 * Rnd) & ".CATDrawing"
End If


Set oActDoc = CATIA.Documents.NewFrom(sTemplateFile)
MsgBox "零件名称:" & vbCrLf & sPartName & vbCrLf & _
        "零件编号:" & vbCrLf & sPartNumber & vbCrLf & _
        "材料:" & vbCrLf & sMaterial & vbCrLf & _
        "质量:" & vbCrLf & sMass

Dim sScale
sScale = ""
sScale = CStr(Round(oActDoc.Sheets.ActiveSheet.Scale2, 2)) & ":1"

Dim oBackGroundView, otxt, n
For n = 1 To oActDoc.Sheets.ActiveSheet.Views.Count
If oActDoc.Sheets.ActiveSheet.Views.Item(n).ViewType = 0 Then
Set oBackGroundView = oActDoc.Sheets.ActiveSheet.Views.Item(n)
End If
Next
For Each otxt In oBackGroundView.Texts
    Select Case otxt.Name
        Case "TitleName"
            otxt.Text = sPartName
        Case "TitlePN"
            otxt.Text = sPartNumber
        Case "TitleMaterial"
            otxt.Text = sMaterial
        Case "TitleScale"
            otxt.Text = sScale
        Case "TitleMass"
            otxt.Text = sMass
        Case Else
    End Select
Next
On Error Resume Next
oActDoc.SaveAs sDrwFileName

End Sub
Sub RefreshTitleBlock(oDocDrw)
IntCATIA
If TypeName(CATIA.ActiveDocument) <> "DrawingDocument" Then
Exit Sub
End If

Set oActDoc = CATIA.ActiveDocument

Dim oBackGroundView, otxt, n
For n = 1 To oActDoc.Sheets.ActiveSheet.Views.Count
If oActDoc.Sheets.ActiveSheet.Views.Item(n).ViewType = 0 Then
Set oBackGroundView = oActDoc.Sheets.ActiveSheet.Views.Item(n)
End If
Next

Dim sScale
sScale = ""
sScale = CStr(Round(oActDoc.Sheets.ActiveSheet.Scale2, 2)) & ":1"

On Error Resume Next
Dim oLinkedDoc
Set oLinkedDoc = DwgLinkedDoc(oActDoc)

Dim sPartName, sPartNumber, sMass, sVolume, sMaterial
sPartName = ""
sPartNumber = ""
sMass = ""
sVolume = ""
sPartName = sPartName & oLinkedDoc.Product.Definition
sPartNumber = sPartNumber & oLinkedDoc.Product.PartNumber
'sMaterial = oLinkedDoc.Product.Nomenclature  'v021s增加
sMaterial = getMaterial(oLinkedDoc.Product)
sMass = CStr(Format(Round(oLinkedDoc.Product.Analyze.Mass, 3), "0.000")) & "kg"

sVolume = CStr(oLinkedDoc.Product.Analyze.Volume)
'On Error GoTo 0

'MsgBox sPartName & vbCrLf & sPartNumber & vbCrLf & "质量:" & vbCrLf & sMass '& vbCrLf & "体积:" & sVolume
'MsgBox sPartName & vbCrLf & sPartNumber & vbCrLf & "材料:" & vbCrLf & sMaterial & vbCrLf & "质量:" & vbCrLf & sMass '& vbCrLf & "体积:" & sVolume
MsgBox "零件名称:" & vbCrLf & sPartName & vbCrLf & _
        "零件编号:" & vbCrLf & sPartNumber & vbCrLf & _
        "材料:" & vbCrLf & sMaterial & vbCrLf & _
        "质量:" & vbCrLf & sMass


For Each otxt In oBackGroundView.Texts
    Select Case otxt.Name
        Case "TitleName"
            otxt.Text = sPartName
        Case "TitlePN"
            otxt.Text = sPartNumber
        Case "TitleMaterial"
            otxt.Text = sMaterial
        Case "TitleScale"
            otxt.Text = sScale
        Case "TitleMass"
            otxt.Text = sMass
        Case Else
    End Select
Next

End Sub

Private Function getMaterial(oProduct1)

getMaterial = ""

If opMat1.Value = True Then
    getMaterial = oProduct1.Nomenclature
Else
    Dim parameters1 ' As Parameters
    Set parameters1 = oProduct1.Parameters
    Dim i As Integer
    Dim matparacn, matparaen
    matparacn = oProduct1.PartNumber & "\" & "材料"
    matparaen = oProduct1.PartNumber & "\" & "Material"
    For i = 1 To parameters1.Count
    If (parameters1.Item(i).Name = matparacn) Or (parameters1.Item(i).Name = matparaen) Then
    getMaterial = parameters1.Item(i).ValueAsString
    End If
    Next
End If
End Function


Private Sub ReadConf()
Dim sConfpath As String
sConfpath = oCATVBA_Folder.Path
Dim objFSO
Set objFSO = CreateObject("Scripting.FileSystemObject")
'MsgBox sConfpath
'MsgBox Dir(sConfpath)
If Not objFSO.FileExists(sConfpath & "\Conf1.ini") Then
Exit Sub
End If
'---读取配置文件---
'On Error Resume Next
    Dim read_OK ' As Long
    Dim read2 As String
    read2 = String(255, 0)
    read_OK = GetPrivateProfileString("新建图纸", "术语（Norm）用作材料", "True", read2, 256, sConfpath & "\Conf1.ini")
    opMat1.Value = read2
    read2 = String(255, 0)
    read_OK = GetPrivateProfileString("新建图纸", "使用模型树中的材料名", "False", read2, 256, sConfpath & "\Conf1.ini")
    opMat2.Value = read2
End Sub
Private Sub SaveConf()
On Error Resume Next
Dim sConfpath As String
Dim write1 'As Long
    '参数一： Section Name (节的名称)。
    '参数二： 节下面的项目名称。
    '参数三： 项目的内容。
    '参数四： ini配置文件的路径名称。
              
sConfpath = oCATVBA_Folder.Path
write1 = WritePrivateProfileString("新建图纸", "术语（Norm）用作材料", CStr(opMat1.Value), sConfpath & "\Conf1.ini")
write1 = WritePrivateProfileString("新建图纸", "使用模型树中的材料名", CStr(opMat2.Value), sConfpath & "\Conf1.ini")
End Sub

Private Sub UserForm_Activate()
ReadConf
End Sub

Private Sub UserForm_Terminate()
SaveConf
End Sub


Public Sub buildPopForm(ByVal a As Object, Optional cp As String = "Properties")
'显示标题栏数据待用户确认或修改
'输入为object, 字典
If TypeName(a) <> "Dictionary" Then
Exit Sub
End If

Dim ak As Variant
ak = a.keys

Dim i, L, T
Dim crl As Object
'重设弹出的窗口尺寸
TitleBlockProperties.Height = 90
TitleBlockProperties.Width = 350
TitleBlockProperties.btnOK.Left = 5
TitleBlockProperties.btnOK.Width = 330
TitleBlockProperties.btnOK.Height = 24

For i = 0 To UBound(ak)
    Set crl = TitleBlockProperties.Frame1.Controls.Add("Forms.Label.1", ak(i), True)
    With crl
        .Caption = ak(i) & " :"
        .Left = 10
        .Top = 10 + 20 * i
        .Width = TitleBlockProperties.Frame1.Width / 3 - .Left - 10
        .TextAlign = 3
    End With
        TitleBlockProperties.Height = TitleBlockProperties.Height + 20
        TitleBlockProperties.btnOK.Top = TitleBlockProperties.Height - 55

    
    Set crl = TitleBlockProperties.Frame1.Controls.Add("Forms.TextBox.1", a(ak(i))(0), True)
    With crl
        .Text = a(ak(i))(1)
        .Left = 10 + 100
        .Top = 10 + 20 * i
        .Width = (TitleBlockProperties.Frame1.Width / 3) * 2 - 20
    End With
Next
TitleBlockProperties.Frame1.ScrollHeight = 20 * (UBound(ak) + 2)
If TitleBlockProperties.Height > 400 Then
    TitleBlockProperties.Height = 400
End If
TitleBlockProperties.Frame1.Height = TitleBlockProperties.Height - 60
TitleBlockProperties.btnOK.Top = TitleBlockProperties.Height - 55
    
TitleBlockProperties.Caption = cp
TitleBlockProperties.Show

End Sub

Public Sub AddDoc2Tree(sDocTemplate As String) 'doctype must be CATPart or CATProduct
Set oActDoc = Nothing
IntCATIA
Dim newPrdDoc As Document
Dim saveFolder As String
saveFolder = "D:"
Dim saveFullName As String

Dim doctype As String
doctype = ""
If Right(sDocTemplate, 8) = ".CATPart" Then
   doctype = ".CATPart"
ElseIf Right(sDocTemplate, 11) = ".CATProduct" Then
   doctype = ".CATProduct"
ElseIf Right(sDocTemplate, 11) = ".CATDrawing" Then
   doctype = ".CATDrawing"
End If
If doctype = "" Then
MsgBox "模板类型错误！", vbInformation, "臭豆腐工具箱CATIA版"
Exit Sub
End If

If oActDoc Is Nothing Then
Call newDocFrom(sDocTemplate)
Exit Sub
End If

saveFolder = oActDoc.Path

Select Case TypeName(oActDoc)
    Case "ProductDocument"
        On Error Resume Next
        Set oSel = Sel("Product")
        If Err.Number <> 0 Then
        Exit Sub
        Else
            
            If ProductIsComponent(oSel) Then
            saveFolder = oSel.Parent.Parent.ReferenceProduct.Parent.Path
            Else
            saveFolder = oSel.ReferenceProduct.Parent.Path
            End If
            
            If saveFolder = "" Then
            MsgBox "找不到当前所选装配的保存路径，请先保存当前再添加子件！", vbInformation + vbOKOnly, "CDF Toolkit for CATIA"
            Exit Sub
            End If
            
            Set newPrdDoc = newDocFrom(sDocTemplate)
             If Err.Number <> 0 Then
             Exit Sub
             End If
            Err.Clear
            
            If doctype <> ".CATDrawing" Then
                saveFullName = InputBox("保存文件:", "CDF Toolkit for CATIA || 唐国庆 13507468076", saveFolder & "\" & newPrdDoc.Product.PartNumber & doctype)
                If saveFullName <> "" Then
                    newPrdDoc.SaveAs (saveFullName)
                    If Err.Number <> 0 Then
                    MsgBox "保存时出了问题，请检查文件路径是否完整或文件是否已重名？你需要手动另存或添加！", vbInformation + vbOKOnly, "CDF Toolkit for CATIA"
                    Exit Sub
                    End If
                Else
                    MsgBox "文件未保存,你需要手动另存或添加！", vbInformation + vbOKOnly, "CDF Toolkit for CATIA"
                    Exit Sub
                End If
                Err.Clear
                Dim ifilelist(1)
                ifilelist(0) = newPrdDoc.FullName
                Dim DesInst As String
                DesInst = newPrdDoc.Product.Definition
                newPrdDoc.Close
                oSel.Products.AddComponentsFromFiles ifilelist, "All"
                If Err.Number <> 0 Then
                   MsgBox "无法在选定零件下添加子件！" & vbCrLf & "零件/产品已经创建，但未成功添加到当前目录树！", vbInformation + vbOKOnly, "CDF Toolkit for CATIA"
                Else
                    If DesInst <> "" Then
                    oSel.Products.Item(oSel.Products.Count).Name = DesInst
                    End If
                End If
            Else
                saveFullName = InputBox("保存文件:", "CDF Toolkit for CATIA || 唐国庆 13507468076", saveFolder & "\" & oSel.PartNumber & doctype)
                If saveFullName <> "" Then
                    newPrdDoc.SaveAs (saveFullName)
                    If Err.Number <> 0 Then
                    MsgBox "保存时出了问题，请检查文件路径是否完整或文件是否已重名？你需要手动另存或添加！", vbInformation + vbOKOnly, "CDF Toolkit for CATIA"
                    Exit Sub
                    End If
                Else
                    MsgBox "文件未保存,你需要手动另存或添加！", vbInformation + vbOKOnly, "CDF Toolkit for CATIA"
                    Exit Sub
                End If
                Err.Clear
            End If
        End If
    Case Else
         Call newDocFrom(sDocTemplate)
End Select

End Sub


Private Function newDocFrom(templatefullpath As String) As Document
On Error Resume Next
Set newDocFrom = CATIA.Documents.NewFrom(templatefullpath)
If Err.Number <> 0 Then
MsgBox "从模板新建失败，可能的原因是:" & vbCrLf & _
       "1. 臭豆腐工具箱CATIA版没有按指导安装，本插件需要安装在特定的目录下 C:或D:\UGmeetsCATIA\.." & vbCrLf & _
       "2. 模板文件名包含中文或特殊字符" & vbCrLf & _
       "3. 当前使用的CATIA版本比模板文件的CATIA版本低,CATIA打不开模板文件,模板文件应当放在" & vbCrLf & _
       oCATVBA_Folder("Template\Start_Part"), vbInformation, "臭豆腐工具箱CATIA版"
Exit Function
End If
If Right(templatefullpath, 11) = ".CATDrawing" Then
Exit Function
End If

'弹出窗口让用户输入必要的信息
Dim a As Variant
Set a = CreateObject("Scripting.Dictionary")
Dim oPrd
Set oPrd = newDocFrom.Product

Dim i As Integer
Dim s As String
Dim k As String
'固有属性
Dim b As Variant
b = Array("PartNumber", "Revision", "Definition", "Nomenclature", "DescriptionRef")
Randomize
For i = 0 To UBound(b)
If b(i) = "PartNumber" Then
a.Add b(i), Array(b(i), Int(100000 + (899999 * Rnd)))
Else
a.Add b(i), Array(b(i), "")
End If
Next

'固有属性

For i = 1 To oPrd.UserRefProperties.Count
    k = ""
    s = ""
    k = GetUserPropStr(oPrd.ReferenceProduct.UserRefProperties.Item(i).Name)
    s = oPrd.UserRefProperties.Item(i).ValueAsString
    'Debug.Print "k=" & k & "(" & "k" & "," & s & ")"
    a.Add k, Array(k, s)
Next

buildPopForm a, "Properties"
'从模板中新建文档

'修改新建文档的属性

End Function
Function ProductIsComponent(iProduct) As Boolean
Dim objDoc As Object
Dim objParentDoc As Object
ProductIsComponent = False
On Error Resume Next
Set objDoc = iProduct.ReferenceProduct.Parent
Set objParentDoc = iProduct.Parent.Parent.ReferenceProduct.Parent
ProductIsComponent = (objDoc Is objParentDoc)

Set objDoc = Nothing
Set objParentDoc = Nothing
Err.Clear

End Function
Function GetUserPropStr(s As String)
GetUserPropStr = Right(s, Len(s) - InStrRev(s, "\"))
End Function
