VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Part2Prod 
   Caption         =   "臭豆腐工具箱CATIA版 | 零件转产品"
   ClientHeight    =   5655
   ClientLeft      =   105
   ClientTop       =   435
   ClientWidth     =   7455
   OleObjectBlob   =   "Part2Prod.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "Part2Prod"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim oOutPutFolder As Object
Dim sOutPutFolder As String
Dim intSN As Long    '生成的零件编号流水号
#If VBA7 Then
Private Declare PtrSafe Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As LongPtr
Private Declare PtrSafe Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As LongPtr
#Else
Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Private Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long
#End If




Private Sub cmdBOM_Click()
'
    IntCATIA
    On Error Resume Next
    If TypeName(oActDoc) <> "PartDocument" Then Exit Sub
    Dim InputCheck As Boolean
    InputCheck = CheckInput()
    If InputCheck = False Then
    Exit Sub
    End If
    'On Error GoTo 0

    txtPreFix.Enabled = False
    txtSerial.Enabled = False
    txtSurfix.Enabled = False
    txtStep.Enabled = False

    CATIA.DisplayFileAlerts = False



'***********主程序*********
Dim BOM(1000, 8)
'名称,类型，状态，表面积，材料，密度，体积，质量
Dim Row As Integer '处理的第几个，行数
Row = 0
BOM(Row, 0) = "名称"
BOM(Row, 1) = "类型"
BOM(Row, 2) = "状态"
BOM(Row, 3) = "表面积"
BOM(Row, 4) = "材料"
BOM(Row, 5) = "密度"
BOM(Row, 6) = "体积"
BOM(Row, 7) = "质量"

Dim objSPAWkb
Set objSPAWkb = CATIA.ActiveDocument.GetWorkbench("SPAWorkbench")


Dim oBody As Object

Dim objRef As Object
Dim objMeasurable
Dim relt, relti 'relations
Dim oManager 'As MaterialManager
Set oManager = oActDoc.Part.GetItem("CATMatManagerVBExt")
Dim oAppliedMaterial 'As Material

        For Each oBody In oActDoc.Part.Bodies
            If oBody.InBooleanOperation = False Then
                    Row = Row + 1
                    Set objRef = oActDoc.Part.CreateReferenceFromObject(oBody)
                    Set objMeasurable = objSPAWkb.GetMeasurable(objRef)
                    oManager.GetMaterialOnBody oBody, oAppliedMaterial
                        If Not (oAppliedMaterial Is Nothing) Then
                            BOM(Row, 4) = oAppliedMaterial.Name
                            BOM(Row, 5) = oAppliedMaterial.AnalysisMaterial.GetValue("SAMDensity") '密度,default unit kg/m3
                        Else
                            BOM(Row, 4) = "未知"
                            BOM(Row, 5) = 1000
                        End If
                    On Error Resume Next
                    BOM(Row, 0) = oBody.Name
                    BOM(Row, 1) = TypeName(oBody)
                    BOM(Row, 2) = BodyState(oBody)
                    BOM(Row, 3) = objMeasurable.Area    'default unit is m2
                    BOM(Row, 6) = objMeasurable.Volume  'default unit is m3

                    BOM(Row, 7) = Round(BOM(Row, 5) * BOM(Row, 6) * 1000, 3) 'default is kg, transfered to g
                    BOM(Row, 3) = Round(BOM(Row, 3) * 10000, 3) 'transfer to cm2
                    BOM(Row, 6) = Round(BOM(Row, 6) * 1000000, 3) 'transfer to cm3
                    BOM(Row, 5) = Round(BOM(Row, 5) / 1000, 3) 'transfer to g/cm3
                    Set oAppliedMaterial = Nothing
                    On Error GoTo 0
            End If

        Next
        For Each oBody In oActDoc.Part.HybridBodies
            Row = Row + 1
                    On Error Resume Next
                    Row = Row + 1
                    Set objRef = oActDoc.Part.CreateReferenceFromObject(oBody)
                    Set objMeasurable = objSPAWkb.GetMeasurable(objRef)
                    oManager.GetMaterialOnHybridBody oBody, oAppliedMaterial
                        If Not (oAppliedMaterial Is Nothing) Then
                            BOM(Row, 4) = oAppliedMaterial.Name
                            BOM(Row, 5) = oAppliedMaterial.AnalysisMaterial.GetValue("SAMDensity") '密度,default unit kg/m3
                        Else
                            BOM(Row, 4) = "未知"
                            BOM(Row, 5) = 1000
                        End If
                    'On Error Resume Next
                    BOM(Row, 0) = oBody.Name
                    BOM(Row, 1) = TypeName(oBody)
                    BOM(Row, 2) = BodyState(oBody)
                    BOM(Row, 3) = objMeasurable.Area    'default unit is m2
                    BOM(Row, 6) = objMeasurable.Volume  'default unit is m3

                    BOM(Row, 7) = Round(BOM(Row, 5) * BOM(Row, 6) * 1000, 3) 'default is kg, transfered to g
                    BOM(Row, 3) = Round(BOM(Row, 3) * 10000, 3) 'transfer to cm2
                    BOM(Row, 6) = Round(BOM(Row, 6) * 1000000, 3) 'transfer to cm3
                    BOM(Row, 5) = Round(BOM(Row, 5) / 1000, 3) 'transfer to g/cm3
                    Set oAppliedMaterial = Nothing
                    On Error GoTo 0


        Next
        For Each oBody In oActDoc.Part.OrderedGeometricalSets
            Row = Row + 1
                    On Error Resume Next
                    Row = Row + 1
                    Set objRef = oActDoc.Part.CreateReferenceFromObject(oBody)
                    Set objMeasurable = objSPAWkb.GetMeasurable(objRef)
'                    oManager.GetMaterialOnHybridBody oBody, oAppliedMaterial
'                        If Not (oAppliedMaterial Is Nothing) Then
'                            BOM(Row, 4) = oAppliedMaterial.Name
'                            BOM(Row, 5) = oAppliedMaterial.AnalysisMaterial.GetValue("SAMDensity") '密度,default unit kg/m3
'                        Else
'                            BOM(Row, 4) = "未知"
'                            BOM(Row, 5) = 1000
'                        End If
                        
                        Dim part1 'As Part
                        Set part1 = oActDoc.Part
                        
                        Dim parameters1 'As Parameters
                        Set parameters1 = part1.Parameters.sublist(oBody, 1)

                        Dim par
                        Dim hasMtl As Boolean
                        hasMtl = False  '是否包含材料
                        For Each par In parameters1
                        Debug.Print par.Name
                            If InStrRev(par.Name, "Material") <> 0 Or InStrRev(par.Name, "材料") <> 0 Then
                                hasMtl = True
                                Exit For
                            End If
                        Next
                        If hasMtl = True Then
                        BOM(Row, 4) = par.ValueAsString
                        End If
                    'On Error Resume Next
                    BOM(Row, 0) = oBody.Name
                    BOM(Row, 1) = TypeName(oBody)
                    BOM(Row, 2) = BodyState(oBody)
                    BOM(Row, 3) = objMeasurable.Area    'default unit is m2
                    BOM(Row, 6) = objMeasurable.Volume  'default unit is m3

                    BOM(Row, 7) = Round(BOM(Row, 5) * BOM(Row, 6) * 1000, 3) 'default is kg, transfered to g
                    BOM(Row, 3) = Round(BOM(Row, 3) * 10000, 3) 'transfer to cm2
                    BOM(Row, 6) = Round(BOM(Row, 6) * 1000000, 3) 'transfer to cm3
                    BOM(Row, 5) = Round(BOM(Row, 5) / 1000, 3) 'transfer to g/cm3
'                    Set oAppliedMaterial = Nothing
                    Set par = Nothing
                    On Error GoTo 0

        Next

On Error Resume Next
Dim Excel As Object
Set Excel = GetObject(, "EXCEL.Application")

'Before launching excel, make sure it's closed.


If Err.Number <> 0 Then
            Err.Clear
            Set Excel = CreateObject("Excel.Application")
Else
            Err.Clear
            MsgBox "如果Excel正在使用中，请先关闭所有Excel程序！", vbCritical
            Exit Sub
End If


' Declare all of our objects for excel
Dim workbooks 'As workbooks
Dim workbook 'As workbook
Dim Sheets As Object
Dim Sheet As Object
Dim worksheet 'As Excel.worksheet
Dim myworkbook 'As Excel.workbook
Dim myworksheet 'As Excel.worksheet

Set workbooks = Excel.Applcation.workbooks
Set myworkbook = Excel.workbooks.Add
Set myworksheet = Excel.ActiveWorkbook.Add

Excel.Visible = True
Excel.cells(1, 1).Resize(UBound(BOM) + 1, UBound(BOM, 2) + 1) = BOM
Excel.cells.Columns.AutoFit
Excel.cells.Rows.AutoFit
Excel.Range("A1:H1").AutoFilter
'Excel.Range("A1:H1").AutoFilter
'***********主程序*********
    CATIA.Visible = True
    CATIA.DisplayFileAlerts = True
    CATIA.RefreshDisplay = True
    txtPreFix.Enabled = True
    txtSerial.Enabled = True
    txtSurfix.Enabled = True
    txtStep.Enabled = True
    Set Excel = Nothing
'MsgBox "导出完成！"


'
'Dim partDoc As Object
'Set partDoc = CATIA.ActiveDocument.Part
'Dim docname As String
'docname = partDoc.Name
'
'Dim objSPAWkb
'Set objSPAWkb = CATIA.ActiveDocument.GetWorkbench("SPAWorkbench")
'

'Dim i As Integer
'Dim bodyNumber As Integer
'bodyNumber = partDoc.Bodies.Count
'
'Dim RwNum As Integer
'RwNum = 1
'' Dump data into spreadsheet

''Format the columns
'Excel.Range("A:A").ColumnWidth = 5
'Excel.Range("B:B").ColumnWidth = 30
'Excel.Range("C:L").ColumnWidth = 15
'Excel.Range("A:L").Font.Name = "Arial"
'Excel.Range("A:L").Font.Size = 10
'
''Format the cells of the top row
'Excel.Range("1:1").Font.Bold = True
'Excel.Range("1:1").RowHeight = 20
'Excel.Range("1:1").Font.Size = 11
'
''Row one
'Excel.cells(1, 1) = "Number"
'Excel.cells(1, 2) = "Part Body Name"
'Excel.cells(1, 3) = "Body Area"
'Excel.cells(1, 4) = "Body Area / 2"
'Excel.cells(1, 6) = "CoG X"
'Excel.cells(1, 7) = "CoG Y"
'Excel.cells(1, 8) = "CoG Z"
'Excel.cells(1, 10) = "Ixx"
'Excel.cells(1, 11) = "Ixy"
'Excel.cells(1, 12) = "Ixz"
'
'For i = 1 To bodyNumber
'
'            Dim body1 'As Body
'            Set body1 = partDoc.Bodies.Item(i)
'
'            Dim namebody As String
'            namebody = body1.Name
'
'            Dim objRef As Object
'            Dim objMeasurable
'
'            Set objRef = partDoc.CreateReferenceFromObject(body1)
'            Set objMeasurable = objSPAWkb.GetMeasurable(objRef)
'
'            Dim TheInertiasList 'As Inertias
'            Set TheInertiasList = objSPAWkb.Inertias
'
'            Dim NewInertia
'            Set NewInertia = TheInertiasList.Add(body1)
'
'            Dim Matrix(8)
'            NewInertia.GetInertiaMatrix Matrix
'
'            Dim bodyArea
'            bodyArea = objMeasurable.Area
'
'            bodyArea = bodyArea * 1550
'
'            Dim bodyArea2 As Integer
'            bodyArea2 = bodyArea / 2
'
'            Dim objCOG(2)
'            objMeasurable.GetCOG objCOG
'
'            ' Row two
'            Excel.cells(RwNum + 1, 1) = i
'            Excel.cells(RwNum + 1, 2) = namebody
'            Excel.cells(RwNum + 1, 3) = bodyArea
'            Excel.cells(RwNum + 1, 4) = bodyArea2
'            Excel.cells(RwNum + 1, 6) = (objCOG(0) / 25.4)
'            Excel.cells(RwNum + 1, 7) = (objCOG(1) / 25.4)
'            Excel.cells(RwNum + 1, 8) = (objCOG(2) / 25.4)
'            Excel.cells(RwNum + 1, 10) = (Matrix(0) * 3417.171898209)
'            Excel.cells(RwNum + 1, 11) = (Matrix(4) * 3417.171898209)
'            Excel.cells(RwNum + 1, 12) = (Matrix(8) * 3417.171898209)
'
'            RwNum = RwNum + 1
'
'Next 'i
End Sub

Private Sub cmdMutiSel_Click()
    IntCATIA
    If TypeName(oActDoc) <> "PartDocument" Then Exit Sub
    Dim bName As String
    bName = InputBox("请输入名称:", "臭豆腐工具箱CATIA版")
    If bName = "" Then Exit Sub
Dim s2, InputObjectType(), Status, FlagFinish, i
Set s2 = CATIA.ActiveDocument.Selection
s2.Clear
FlagFinish = 0

        ReDim InputObjectType(2)
        InputObjectType(0) = "Body"
        InputObjectType(1) = "HybridBody"
        InputObjectType(2) = "GeometricElement"
On Error Resume Next
Do Until FlagFinish = 1

        Status = s2.SelectElement3(InputObjectType, "Select the " & InputObjectType(0) & ". Press Esc to finish the selection.", 0, 1, 0)
        s2.Item2(1).Value.Name = bName
        If (Status = "Cancel") Then
            FlagFinish = 1
            If MsgBox("将退出点选修改模式，要保存修改么?", vbYesNo, "几何体改同名") = vbYes Then
            oActDoc.Save
            s2.Clear
            End If
        End If
        s2.Clear
Loop
        
End Sub

Private Sub cmdRun1_Click()
    IntCATIA
    If TypeName(oActDoc) <> "PartDocument" Then Exit Sub
    Dim CapStop, BtmColor
    CapStop = cmdRun1.Caption
    BtmColor = cmdRun1.BackColor
    
    Dim InputCheck As Boolean
    InputCheck = CheckInput()
    If InputCheck = False Then
    Exit Sub
    End If
    'cmdRun1.Enabled = False

    txtPreFix.Enabled = False
    txtSerial.Enabled = False
    txtSurfix.Enabled = False
    txtStep.Enabled = False

    CATIA.DisplayFileAlerts = False
    'CATIA.RefreshDisplay = False
'***设置****
    Dim partInfrastructureSettingAtt1 As Object
    Set partInfrastructureSettingAtt1 = CATIA.SettingControllers.Item("CATMmuPartInfrastructureSettingCtrl")
    Dim OrgSetting() As Variant
    ReDim OrgSetting(3)
    OrgSetting(0) = partInfrastructureSettingAtt1.NewWithPanel
    OrgSetting(1) = partInfrastructureSettingAtt1.HybridDesignMode
    OrgSetting(2) = partInfrastructureSettingAtt1.ColorSynchronizationMode
    OrgSetting(3) = partInfrastructureSettingAtt1.ColorSynchronizationEditability
    
    partInfrastructureSettingAtt1.NewWithPanel = False
    partInfrastructureSettingAtt1.HybridDesignMode = False
    If chkColor.Value = True Then
    partInfrastructureSettingAtt1.ColorSynchronizationMode = True
    partInfrastructureSettingAtt1.ColorSynchronizationEditability = True
    Else
    partInfrastructureSettingAtt1.ColorSynchronizationMode = False
    partInfrastructureSettingAtt1.ColorSynchronizationEditability = False
    End If
'***设置****
'***********主程序*********
intSN = Val(txtSerial.Text)
Dim oDocProd As ProductDocument
Set oDocProd = CATIA.Documents.Add("Product")
oDocProd.Product.PartNumber = oActDoc.Product.PartNumber
Dim oProds As Products
Set oProds = oDocProd.Product.Products
Dim oBody As Object
Dim oPart1, sDocPart1Nomenclature, PN
Dim oDocPart1 As Object
Dim oTargSel As Object
        For Each oBody In oActDoc.Part.Bodies
            Set oSelection = oActDoc.Selection
            oSelection.Clear
            If oBody.InBooleanOperation = False Then
                    On Error Resume Next
                    oSelection.Add oBody
                    oSelection.Copy
                 
                    
                    'Dim oPart1, sDocPart1Nomenclature, PN
                    sDocPart1Nomenclature = Replace(Replace(Replace(oBody.Name, "\", "_"), "#", "_"), ".", "_")
                    PN = txtPreFix.Text & PreZero(CStr(intSN), Len(txtSerial.Text)) & txtSurfix.Text
                    If cmdRun1.BackColor = BtmColor Then
                        cmdRun1.BackColor = &HFF80FF
                    Else
                        cmdRun1.BackColor = BtmColor
                    End If
                    cmdRun1.Caption = PN
                    
                    Set oPart1 = oProds.AddNewComponent("Part", txtPreFix.Text & PreZero(CStr(intSN), Len(txtSerial.Text)) & txtSurfix.Text)
                    'Dim oDocPart1 As Object
                    Set oDocPart1 = CATIA.Documents.Item(PN & ".CATPart")
        
                    oSelection.Add oDocPart1.Part
                   
                    oSelection.PasteSpecial "CATPrtResultWithOutLink"
                    
                    oDocPart1.Part.MainBody = oDocPart1.Part.Bodies.Item(oBody.Name)
        
                    
                    oDocPart1.Product.PartNumber = PN
                    oDocPart1.Product.Nomenclature = sDocPart1Nomenclature
                    oDocPart1.Product.DescriptionInst = sDocPart1Nomenclature
                    oDocPart1.Part.Update
                    oPart1.Name = sDocPart1Nomenclature
                    On Error GoTo 0
                    intSN = intSN + Val(txtStep.Text)
                    
            End If
            
        Next

oDocProd.SaveAs txtOutProd.Text



'***********主程序*********
    CATIA.Visible = True
    CATIA.DisplayFileAlerts = True
    CATIA.RefreshDisplay = True
    
    cmdRun1.Caption = CapStop
    cmdRun1.BackColor = BtmColor
    cmdRun1.Enabled = True
    txtPreFix.Enabled = True
    txtSerial.Enabled = True
    txtSurfix.Enabled = True
    txtStep.Enabled = True
    partInfrastructureSettingAtt1.NewWithPanel = OrgSetting(0)
    partInfrastructureSettingAtt1.HybridDesignMode = OrgSetting(1)
    partInfrastructureSettingAtt1.ColorSynchronizationMode = OrgSetting(2)
    partInfrastructureSettingAtt1.ColorSynchronizationEditability = OrgSetting(3)
End Sub

Private Sub cmdRun2_Click()
    IntCATIA
    If TypeName(oActDoc) <> "PartDocument" Then Exit Sub
    Dim InputCheck As Boolean
    InputCheck = CheckInput()
    If InputCheck = False Then
    Exit Sub
    End If

    txtPreFix.Enabled = False
    txtSerial.Enabled = False
    txtSurfix.Enabled = False
    txtStep.Enabled = False

    CATIA.DisplayFileAlerts = False
    Dim keepOrigName As Boolean
    If (MsgBox("是否需要保留原来的几何体名称？" & vbCrLf & _
                        "【是】，不删除原来的几何体名称，只在现有名称之前加上新编号" & vbCrLf & _
                        "【否】，删除原来的几何体名称，仅以新编号作为名称", vbYesNo, "臭豆腐工具箱CATIA版") = vbYes) Then
                        keepOrigName = True
                        Else
                        keepOrigName = False
    End If
                    


'***********主程序*********
intSN = Val(txtSerial.Text)
Dim oBody As Object
Dim PN
        For Each oBody In oActDoc.Part.Bodies
            If oBody.InBooleanOperation = False Then
                    On Error Resume Next
                    PN = txtPreFix.Text & PreZero(CStr(intSN), Len(txtSerial.Text)) & txtSurfix.Text
                    If keepOrigName = True Then
                    PN = PN & "_" & oBody.Name
                    End If
                    oBody.Name = PN

                    On Error GoTo 0
                    intSN = intSN + Val(txtStep.Text)
                    
            End If
            
        Next
        For Each oBody In oActDoc.Part.HybridBodies
                    On Error Resume Next
                    PN = txtPreFix.Text & PreZero(CStr(intSN), Len(txtSerial.Text)) & txtSurfix.Text
                    PN = txtPreFix.Text & PreZero(CStr(intSN), Len(txtSerial.Text)) & txtSurfix.Text
                    If keepOrigName = True Then
                    PN = PN & "_" & oBody.Name
                    End If
                    oBody.Name = PN
                    On Error GoTo 0
                    intSN = intSN + Val(txtStep.Text)
        Next
        For Each oBody In oActDoc.Part.OrderedGeometricalSets
                    On Error Resume Next
                    PN = txtPreFix.Text & PreZero(CStr(intSN), Len(txtSerial.Text)) & txtSurfix.Text
                    PN = txtPreFix.Text & PreZero(CStr(intSN), Len(txtSerial.Text)) & txtSurfix.Text
                    If keepOrigName = True Then
                    PN = PN & "_" & oBody.Name
                    End If
                    oBody.Name = PN
                    On Error GoTo 0
                    intSN = intSN + Val(txtStep.Text)
        Next


'***********主程序*********
    CATIA.Visible = True
    CATIA.DisplayFileAlerts = True
    CATIA.RefreshDisplay = True
    txtPreFix.Enabled = True
    txtSerial.Enabled = True
    txtSurfix.Enabled = True
    txtStep.Enabled = True
MsgBox "请记得 Ctrl+S 保存修改！"
End Sub

Private Sub cmdShortName_Click()
    IntCATIA
    If TypeName(oActDoc) <> "PartDocument" Then Exit Sub
    Dim InputCheck As Boolean
    InputCheck = CheckInput()
    If InputCheck = False Then
    Exit Sub
    End If

    txtPreFix.Enabled = False
    txtSerial.Enabled = False
    txtSurfix.Enabled = False
    txtStep.Enabled = False

    CATIA.DisplayFileAlerts = False

'***********主程序*********
Dim oBody As Object
Dim PN
        For Each oBody In oActDoc.Part.Bodies
            If oBody.InBooleanOperation = False Then
                PN = oBody.Name
                On Error Resume Next
                        If InStrRev(PN, "\") > 0 Then
                        PN = Left(PN, InStrRev(PN, "\") - 1)
                        End If
                        If InStrRev(PN, "\") > 0 Then
                        PN = Right(PN, Len(PN) - InStrRev(PN, "\"))
                        End If
                oBody.Name = PN
                On Error GoTo 0
                    
            End If
            
        Next
        For Each oBody In oActDoc.Part.HybridBodies
                PN = oBody.Name
                On Error Resume Next
                        If InStrRev(PN, "\") > 0 Then
                        PN = Left(PN, InStrRev(PN, "\") - 1)
                        End If
                        If InStrRev(PN, "\") > 0 Then
                        PN = Right(PN, Len(PN) - InStrRev(PN, "\"))
                        End If
                oBody.Name = PN
                On Error GoTo 0
        Next
        For Each oBody In oActDoc.Part.OrderedGeometricalSets
                PN = oBody.Name
                On Error Resume Next
                        If InStrRev(PN, "\") > 0 Then
                        PN = Left(PN, InStrRev(PN, "\") - 1)
                        End If
                        If InStrRev(PN, "\") > 0 Then
                        PN = Right(PN, Len(PN) - InStrRev(PN, "\"))
                        End If
                oBody.Name = PN
                On Error GoTo 0
        Next


'***********主程序*********
    CATIA.Visible = True
    CATIA.DisplayFileAlerts = True
    CATIA.RefreshDisplay = True
    txtPreFix.Enabled = True
    txtSerial.Enabled = True
    txtSurfix.Enabled = True
    txtStep.Enabled = True
MsgBox "请记得 Ctrl+S 保存修改！"
End Sub

Private Sub txtInput_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
Set oActDoc = Nothing
IntCATIA
If oActDoc Is Nothing Then
chkColor.Value = False
txtInput.Text = "当前活动文档不是CATPart类型！" & vbCrLf & "切换到【零件设计】工作台后双击此处重新检测！"
txtOutProd.Text = ""
End If
Select Case TypeName(oActDoc)
    Case "PartDocument"
        Set oOutPutFolder = sTargetFolder(Replace(oActDoc.FullName, ".CATPart", ""))
        sOutPutFolder = oOutPutFolder.Path
        'Debug.Print sOutPutFolder
        txtInput.Text = oActDoc.FullName
        txtOutProd.Text = sOutPutFolder & "\" & Replace(oActDoc.Name, ".CATPart", ".CATProduct")
        lblWarning2.Caption = "- 检测到几何图形集（" & oActDoc.Part.HybridBodies.Count & "）个，有序几何图形集（" & oActDoc.Part.OrderedGeometricalSets.Count & "）个"
    Case Else
        txtInput.Text = "当前活动文档不是CATPart类型！" & vbCrLf & "切换到【零件设计】工作台后双击此处重新检测！"
        txtOutProd.Text = ""
        lblWarning2.Caption = ""
End Select
End Sub


Private Sub UserForm_Activate()
Set oActDoc = Nothing
IntCATIA
If oActDoc Is Nothing Then
txtInput.Text = "当前活动文档不是CATPart类型！" & vbCrLf & "切换到【零件设计】工作台后双击此处重新检测！"
txtOutProd.Text = ""
lblWarning2.Caption = ""
End If

Select Case TypeName(oActDoc)
    Case "PartDocument"
        Set oOutPutFolder = sTargetFolder(Replace(oActDoc.FullName, ".CATPart", ""))
        sOutPutFolder = oOutPutFolder.Path
        'Debug.Print sOutPutFolder
        txtInput.Text = oActDoc.FullName
        txtOutProd.Text = sOutPutFolder & "\" & Replace(oActDoc.Name, ".CATPart", ".CATProduct")
        lblWarning2.Caption = "- 检测到几何图形集（" & oActDoc.Part.HybridBodies.Count & "）个，有序几何图形集（" & oActDoc.Part.OrderedGeometricalSets.Count & "）个"
    Case Else
        txtInput.Text = "当前活动文档不是CATPart类型！" & vbCrLf & "切换到【零件设计】工作台后双击此处重新检测！"
        txtOutProd.Text = ""
        lblWarning2.Caption = ""
End Select
ReadConf
End Sub
Function sTargetFolder(Optional ByVal sFolder As String) As Object
Dim oFSO As Object
   Set oFSO = CreateObject("Scripting.FileSystemObject")
   If Not oFSO.FolderExists(sFolder) Then
    oFSO.CreateFolder (sFolder)
   End If
   
Set sTargetFolder = oFSO.GetFolder(sFolder)
End Function


Private Sub txtOutProd_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
On Error Resume Next
Shell "explorer.exe " & sOutPutFolder, vbNormalFocus
On Error GoTo 0
End Sub
Private Function TrimedStr(ByVal s As String, ByVal n As Integer, Optional ByVal LR As String = "R")
'截取给定字符串的左边或右边的n个字符
If Len(s) > n Then
    If LR = "R" Then
        TrimedStr = ".." & Right(s, n)
    ElseIf LR = "L" Then
        TrimedStr = Left(s, n) & ".."
    End If
Else
        TrimedStr = s
End If
End Function


Private Sub txtSerial_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    Select Case KeyAscii
          Case Asc("0") To Asc("9")
          Case 8
          Case Else
            MsgBox "只能输入数字！"
    End Select
End Sub
Private Sub txtStep_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    Select Case KeyAscii
          Case Asc("0") To Asc("9")
          Case 8
          Case Else
            MsgBox "只能输入数字！"
    End Select
End Sub
Function PreZero(s As String, di As Integer)
If di > 1 And di < 20 Then
    Do While Len(s) < CInt(di)
         s = "0" & CStr(s)
    Loop
End If
PreZero = s
End Function
Function CheckInput() As Boolean
CheckInput = True
    IntCATIA
    If oActDoc Is Nothing Then
        MsgBox "没有找到活动文档，是否同时打开了几个CATIA进程?"
        CheckInput = False
        Exit Function
    ElseIf oActDoc.FullName <> txtInput.Text Then
        MsgBox "当前活动文档与输入栏信息不符，请刷新输入栏！", vbInformation
        CheckInput = False
        Exit Function
    End If
    If IsNumeric(txtStep) = False Then
            MsgBox "步长填写错误"
            CheckInput = False
            Exit Function
        ElseIf Val(txtStep) <= 0 Then
            MsgBox "步长填写错误"
            CheckInput = False
            Exit Function
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
    read_OK = GetPrivateProfileString("零件转产品", "保留颜色", "True", read2, 256, sConfpath & "\Conf1.ini")
    chkColor.Value = read2
    read2 = String(255, 0)
    read_OK = GetPrivateProfileString("零件转产品", "子件编号前缀", "CDF-1688-", read2, 256, sConfpath & "\Conf1.ini")
    txtPreFix.Text = read2
    read2 = String(255, 0)
    read_OK = GetPrivateProfileString("零件转产品", "子件编号后缀", "", read2, 256, sConfpath & "\Conf1.ini")
    txtSurfix.Text = read2
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
write1 = WritePrivateProfileString("零件转产品", "保留颜色", CStr(chkColor.Value), sConfpath & "\Conf1.ini")
write1 = WritePrivateProfileString("零件转产品", "子件编号前缀", txtPreFix.Text, sConfpath & "\Conf1.ini")
write1 = WritePrivateProfileString("零件转产品", "子件编号后缀", txtSurfix.Text, sConfpath & "\Conf1.ini")

End Sub

Private Sub UserForm_Terminate()
SaveConf
End Sub
Function BodyState(ByVal obody9 As Object) As String
BodyState = "显示/Show"

Dim Sel9 As Object

Set Sel9 = oActDoc.Selection
Sel9.Clear

Dim VisPropertyset As Object
Dim showstate
Sel9.Add obody9
Set VisPropertyset = Sel9.VisProperties
VisPropertyset.GetShow showstate
If showstate = 1 Then    '1为隐藏
BodyState = "隐藏/Hide"
End If

Sel9.Clear
Set Sel9 = Nothing
Set obody9 = Nothing
Set VisPropertyset = Nothing

End Function
Function bodyArea(ByVal obody9 As Object) As Double
bodyArea = 0

Dim Sel9 As Object

Set Sel9 = oActDoc.Selection
Sel9.Clear

Sel9.Add obody9
Dim meas1
Set meas1 = oActDoc.GetWorkbench("SPAWorkbench").GetMeasurable(oActDoc.Part.CreateReferenceFromObject(Sel9.Item2(1).Value))
bodyArea = meas1.Volume
Sel9.Clear
Set Sel9 = Nothing
Set obody9 = Nothing
Set VisPropertyset = Nothing

End Function
