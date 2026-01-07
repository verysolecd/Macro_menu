VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Batch_ReName 
   Caption         =   "零件批量重命名"
   ClientHeight    =   11205
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   2310
   OleObjectBlob   =   "Batch_ReName.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "Batch_ReName"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim DicP As Object
Dim DicPAll As Object
Dim UnloadQty As Integer  '未加载的文件数量

#If VBA7 Then
Private Declare PtrSafe Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As LongPtr
Private Declare PtrSafe Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As LongPtr
#Else
Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Private Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long
#End If




Private Sub chkBlankPart_Change()
If chkBlankPart.Value = True Then
txtBlankQty.Enabled = True
Else
txtBlankQty.Enabled = False
End If
End Sub


Private Sub cmdAddPreSuf_Click()
'检查输入有效性
If txtAddPre.Text = "" And txtAddSuf.Text = "" Then
MsgBox "请先输入要添加的前缀或后缀！", vbOKOnly, "臭豆腐工具箱CATIA版"
Exit Sub
End If

IntCATIA
'-----------设置设计模式-------------
CATIA.DisplayFileAlerts = False
'Dim settingControllers1
'Dim cacheSettingAtt1
'Set settingControllers1 = CATIA.SettingControllers
'Set cacheSettingAtt1 = settingControllers1.Item("CATSysCacheSettingCtrl")
'cacheSettingAtt1.ActivationMode = 1
'Set cacheSettingAtt1 = Nothing
'Set settingControllers1 = Nothing
Dim oProdxx
Set oProdxx = oActDoc.Product
oProdxx.ApplyWorkMode DESIGN_MODE
Set oProdxx = Nothing
'-----------设置设计模式-------------
Dim sDocPath As String
sDocPath = oActDoc.Path '获取当前文档保存路径
DicPAll.RemoveAll

GetAllPN oActDoc.Product

On Error Resume Next
Set oSel = Sel("Product")
If Err.Number <> 0 Then
Exit Sub
End If

Dim oMyProducts
Set oMyProducts = oSel.Products
GetTargetDic oSel


If oMyProducts.Count <> 0 Then
    CATIA.DisplayFileAlerts = False

    Dim i As Integer
    Dim oProduct9 As Object

    For Each oProduct9 In oMyProducts                           '##0
       Dim NewPNi As String
       'On Error Resume Next
       NewPNi = txtAddPre.Text & oProduct9.PartNumber & txtAddSuf.Text '第二种批量修改的情况
       If Not DicPAll.exists(NewPNi) Then               '##0 新料号不存在内存中才改，存在则不改
              If Not ProductIsComponent(oProduct9) Then '--1

               If DicP.exists(oProduct9.PartNumber) Then

                    Dim NewName As String
                                    If chkOldFolder.Value = True Then
                                    sDocPath = oProduct9.ReferenceProduct.Parent.Path
'                                    MsgBox sDocPath
                                    End If

'                    NewName = txtAddPre.Text & oProduct9.Name & txtAddSuf.Text
                    NewName = oProduct9.Name
                    oProduct9.PartNumber = NewPNi
                    oProduct9.Name = NewName
                    Err.Clear
                    On Error Resume Next

                    
                    oProduct9.ReferenceProduct.Parent.SaveAs sDocPath & "\" & NewPNi
                    If Err.Number <> 0 Then
                        Dim NewPN2 As String
                        Randomize
                        NewPN2 = NewPNi & "_" & Int(100 * Rnd)
                        MsgBox sDocPath & "\" & NewPNi & " 无法保存,可能已经存在同名文件" & vbCrLf & _
                               "将保存为" & sDocPath & "\" & NewPN2
                        oProduct9.PartNumber = NewPN2
                        oProduct9.ReferenceProduct.Parent.SaveAs sDocPath & "\" & NewPN2
                    End If
                    DicPAll.Add oProduct9.PartNumber, oProduct9.ReferenceProduct.Parent.FullName
                    On Error GoTo 0
               End If
             End If                                              '--1
        End If                              '##0
     On Error GoTo 0
    Next
End If


'记录需要删除的旧文件
If cbxDelOld.Value = True Then
        Dim AArKill
        AArKill = DicP.keys()
    On Error Resume Next
    If cbxDelOld.Value = True And UBound(AArKill) <> -1 Then
        If MsgBox("确实要删除旧文件吗? 删除后无法找回，请谨慎操作", vbYesNo + vbInformation, "零件批量重命名") = vbYes Then
        
            Dim kk As Integer
            For kk = 0 To UBound(AArKill)
            Kill DicP.Item(AArKill(kk))
            Next
        End If
    End If
End If
DicP.RemoveAll
oActDoc.Save
If Err.Number <> 0 Then
MsgBox "请Ctrl + S 保存!", vbOKOnly, "臭豆腐工具箱CATIA版"
End If
CATIA.DisplayFileAlerts = True
End Sub

Private Sub cmdInsName_Click()
cmdinsNameOption 1
End Sub

Private Sub cmdInsName2_Click()
cmdinsNameOption 2
End Sub
Private Sub cmdInsName3_Click()
cmdinsNameOption 3
End Sub
Private Sub cmdInsName4_Click()
cmdinsNameOption 4
End Sub
Private Sub cmdinsNameOption(op As Integer)
cbxIgnoreHide.Enabled = False
IntCATIA
If TypeName(oActDoc) <> "ProductDocument" Then
   If TypeName(oActDoc) <> "PartDocument" Then
        MsgBox "当前文件不是CATIA零件或产品 !" & vbCrLf & "此命令只能在零件或产品文件中运行!"
        Exit Sub
   End If
End If
CATIA.DisplayFileAlerts = False
'Dim settingControllers1
'Dim cacheSettingAtt1
'Set settingControllers1 = CATIA.SettingControllers
'Set cacheSettingAtt1 = settingControllers1.Item("CATSysCacheSettingCtrl")
'cacheSettingAtt1.ActivationMode = 1
'Set cacheSettingAtt1 = Nothing
'Set settingControllers1 = Nothing
Dim oProdxx
Set oProdxx = oActDoc.Product
oProdxx.ApplyWorkMode DESIGN_MODE
Set oProdxx = Nothing
'-----------设置设计模式-------------
Dim s2, InputObjectType(), Status, FlagFinish, i
Set s2 = CATIA.ActiveDocument.Selection
s2.Clear

On Error Resume Next

        ReDim InputObjectType(0)
        InputObjectType(0) = "Product"
Do Until FlagFinish = 1

        Status = s2.SelectElement3(InputObjectType, "Select the " & InputObjectType(0) & ". Press Esc to finish the selection.", 0, 1, 0)

        If (Status = "Cancel") Then
            FlagFinish = 1
            If MsgBox("将退出点选修改模式，要保存修改么?", vbYesNo, "零件批量重命名") = vbYes Then
            oActDoc.Save
            s2.Clear
            cbxIgnoreHide.Enabled = True
            Exit Sub
            End If
        End If
        
        For i = 1 To s2.Count2
         Err.Clear
         Dim MyNewName As String
'         MyNewName = s2.Item2(i).Value.PartNumber
         'MsgBox MyNewName
         Select Case op
                Case 1
                MyNewName = s2.Item2(i).Value.PartNumber
                Case 2
                MyNewName = s2.Item2(i).Value.Definition
                Case 3
                MyNewName = s2.Item2(i).Value.Nomenclature
                Case 4
                MyNewName = s2.Item2(i).Value.DescriptionRef
         End Select
         
         
         
         If Err.Number = 0 Then
           Dim j As Integer
           Dim parentAssy As Object
           Set parentAssy = s2.Item2(i).Value.Parent.Parent
           j = 0
           Err.Clear
            If TypeName(parentAssy) = "Product" Then
                Do
                Err.Clear
                 j = j + 1
                 parentAssy.ReferenceProduct.Products.Item(s2.Item2(i).Value.Name).Name = MyNewName & "." & j
                Loop Until Err.Number = 0 Or j > 500
            End If
            s2.Item2(i).Value.ReferenceProduct.Parent.Save
        End If
        Next
        s2.Clear
Loop
cbxIgnoreHide.Enabled = True
End Sub

Private Sub cmdPick1_Click()
cbxIgnoreHide.Enabled = False

IntCATIA
If TypeName(oActDoc) <> "ProductDocument" Then
   If TypeName(oActDoc) <> "PartDocument" Then
        MsgBox "当前文件不是CATIA零件或产品 !" & vbCrLf & "此命令只能在零件或产品文件中运行!"
        Exit Sub
   End If
End If
CATIA.DisplayFileAlerts = False
'Dim settingControllers1
'Dim cacheSettingAtt1
'Set settingControllers1 = CATIA.SettingControllers
'Set cacheSettingAtt1 = settingControllers1.Item("CATSysCacheSettingCtrl")
'cacheSettingAtt1.ActivationMode = 1
'Set cacheSettingAtt1 = Nothing
'Set settingControllers1 = Nothing
Dim oProdxx
Set oProdxx = oActDoc.Product
oProdxx.ApplyWorkMode DESIGN_MODE
Set oProdxx = Nothing
'-----------设置设计模式-------------
Dim sDocPath As String
sDocPath = oActDoc.Path '获取当前文档保存路径

DicPAll.RemoveAll

GetAllPN oActDoc.Product

Dim Num As Long

'Num = Int(txtStartNum.Text)
Num = CLng(txtStartNum.Text)
Dim s2, InputObjectType(), Status, FlagFinish, i
Set s2 = CATIA.ActiveDocument.Selection
s2.Clear
DicP.RemoveAll
FlagFinish = 0

        ReDim InputObjectType(0)
        InputObjectType(0) = "Product"
Do Until FlagFinish = 1

        Status = s2.SelectElement3(InputObjectType, "Select the " & InputObjectType(0) & ". Press Esc to finish the selection.", 0, 1, 0)

        If (Status = "Cancel") Then
            FlagFinish = 1
            If MsgBox("将退出点选修改模式，要保存修改么?", vbYesNo, "零件批量重命名") = vbYes Then
            oActDoc.Save
            s2.Clear
            cbxIgnoreHide.Enabled = True
            End If
        End If
        
        For i = 1 To s2.Count2
               Dim NewPNi As String
                NewPNi = txtPreFix.Text & PreZero(CStr(Num), Len(txtStartNum.Text)) & txtSufFix.Text '第一种批量修改的情况
                  Do Until Not DicPAll.exists(NewPNi)   '如果新料号已经存在于当前内存，则跳过取下一个
                          If DicP.exists(NewPNi) Then   '与新命名同规则的不需要改名
                          DicP.Remove NewPNi
                          End If
                      Num = Num + Int(txtStep.Text)
                      NewPNi = txtPreFix.Text & PreZero(CStr(Num), Len(txtStartNum.Text)) & txtSufFix.Text '第一种批量修改的情况
                  Loop
        
            If Not ProductIsComponent(s2.Item2(i).Value) Then
               On Error Resume Next
               '记录修改的文件，准备删除
                If Not DicP.exists(s2.Item2(i).Value.PartNumber) Then
                    If Err.Number = 0 Then
                       
                        'MsgBox S2.Item2(i).Value.PartNumber & vbCrLf & S2.Item2(i).Value.ReferenceProduct.Parent.FullName
                        DicP.Add s2.Item2(i).Value.PartNumber, s2.Item2(i).Value.ReferenceProduct.Parent.FullName
                        Dim NewName As String
                        
                        NewName = Replace(s2.Item2(i).Value.Name, s2.Item2(i).Value.PartNumber, NewPNi)
                        If chkOldFolder.Value = True Then
                        sDocPath = s2.Item2(i).Value.ReferenceProduct.Parent.Path
'                                    MsgBox sDocPath
                        End If
                        If Not DicPAll.exists(NewPNi) Then
                        SaveNew s2.Item2(i).Value, NewPNi, NewName, sDocPath
                        End If
                    End If
                End If
                Num = Num + Int(txtStep.Text)
                txtStartNum.Text = PreZero(CStr(Num), Len(txtStartNum.Text))
                Err.Clear
            End If
        Next
    s2.Clear

Loop

'记录需要删除的旧文件
If cbxDelOld.Value = True Then
    Dim AArKill
    AArKill = DicP.keys()

    If cbxDelOld.Value = True And UBound(AArKill) <> -1 Then
        If MsgBox("确实要删除旧文件吗? 删除后无法找回，请谨慎操作", vbYesNo + vbInformation, "零件批量重命名") = vbYes Then
        
            Dim kk As Integer
            For kk = 0 To UBound(AArKill)
            'MsgBox AArKill(kk)
            Kill DicP.Item(AArKill(kk))
            Next
        End If
    End If
End If
DicP.RemoveAll
FlagFinish = 0
cbxIgnoreHide.Enabled = True
Set s2 = Nothing
End Sub

Private Sub cmdPick2_Click()
cbxIgnoreHide.Enabled = False
IntCATIA
If TypeName(oActDoc) <> "ProductDocument" Then
   If TypeName(oActDoc) <> "PartDocument" Then
        MsgBox "当前文件不是CATIA零件或产品 !" & vbCrLf & "此命令只能在零件或产品文件中运行!"
        Exit Sub
   End If
End If
'检查输入有效性
If txtAddPre.Text = "" And txtAddSuf.Text = "" Then
MsgBox "请先输入要添加的前缀或后缀！", vbOKOnly, "臭豆腐工具箱CATIA版"
Exit Sub
End If

Dim sDocPath As String
sDocPath = oActDoc.Path '获取当前文档保存路径


'-----------设置设计模式-------------
CATIA.DisplayFileAlerts = False
'Dim settingControllers1
'Dim cacheSettingAtt1
'Set settingControllers1 = CATIA.SettingControllers
'Set cacheSettingAtt1 = settingControllers1.Item("CATSysCacheSettingCtrl")
'cacheSettingAtt1.ActivationMode = 1
'Set cacheSettingAtt1 = Nothing
'Set settingControllers1 = Nothing
Dim oProdxx
Set oProdxx = oActDoc.Product
oProdxx.ApplyWorkMode DESIGN_MODE
'-----------设置设计模式-------------
On Error Resume Next
DicPAll.RemoveAll

GetAllPN oActDoc.Product

Dim s2, InputObjectType(), Status, FlagFinish, i
Set s2 = CATIA.ActiveDocument.Selection
s2.Clear
DicP.RemoveAll
FlagFinish = 0

        ReDim InputObjectType(0)
        InputObjectType(0) = "Product"
Do Until FlagFinish = 1

        Status = s2.SelectElement3(InputObjectType, "Select the " & InputObjectType(0) & ". Press Esc to finish the selection.", 0, 1, 0)

        If (Status = "Cancel") Then
            FlagFinish = 1
            If MsgBox("将退出点选修改模式，要保存修改么?", vbYesNo, "零件批量重命名") = vbYes Then
            oActDoc.Save
            s2.Clear
            cbxIgnoreHide.Enabled = True
            End If
        End If
        
        For i = 1 To s2.Count2
            If Not ProductIsComponent(s2.Item2(i).Value) Then
               
               '记录修改的文件，准备删除
                If Not DicP.exists(s2.Item2(i).Value.PartNumber) Then
                    If Err.Number = 0 Then
                       
                        'MsgBox S2.Item2(i).Value.PartNumber & vbCrLf & S2.Item2(i).Value.ReferenceProduct.Parent.FullName
                        DicP.Add s2.Item2(i).Value.PartNumber, s2.Item2(i).Value.ReferenceProduct.Parent.FullName
                        Dim NewPN As String
                        Dim NewName As String
                        
                        NewPN = txtAddPre.Text & s2.Item2(i).Value.PartNumber & txtAddSuf.Text
'                        NewName = txtAddPre.Text & S2.Item2(i).Value.Name & txtAddSuf.Text
                        NewName = s2.Item2(i).Value.Name
                            If chkOldFolder.Value = True Then
                            sDocPath = s2.Item2(i).Value.ReferenceProduct.Parent.Path
'                                    MsgBox sDocPath
                            End If
                        
                        If Not DicPAll.exists(NewPN) Then
                        SaveNew s2.Item2(i).Value, NewPN, NewName, sDocPath
                        End If
                    End If
                End If
                Err.Clear
            End If
        Next
    s2.Clear

Loop

'记录需要删除的旧文件
If cbxDelOld.Value = True Then
        Dim AArKill
        AArKill = DicP.keys()
    
    If cbxDelOld.Value = True And UBound(AArKill) <> -1 Then
        If MsgBox("确实要删除旧文件吗? 删除后无法找回，请谨慎操作", vbYesNo + vbInformation, "零件批量重命名") = vbYes Then
        
            Dim kk As Integer
            For kk = 0 To UBound(AArKill)
            'MsgBox AArKill(kk)
            Kill DicP.Item(AArKill(kk))
            Next
        End If
    End If
End If
DicP.RemoveAll
FlagFinish = 0
cbxIgnoreHide.Enabled = True
Set s2 = Nothing
End Sub


Private Sub cmdReName_Click()
'检查输入有效性
IntCATIA
'-----------设置设计模式-------------
CATIA.DisplayFileAlerts = False
'Dim settingControllers1
'Dim cacheSettingAtt1
'Set settingControllers1 = CATIA.SettingControllers
'Set cacheSettingAtt1 = settingControllers1.Item("CATSysCacheSettingCtrl")
'cacheSettingAtt1.ActivationMode = 1
'Set cacheSettingAtt1 = Nothing
'Set settingControllers1 = Nothing
Dim oProdxx
Set oProdxx = oActDoc.Product
oProdxx.ApplyWorkMode DESIGN_MODE
Set oProdxx = Nothing
'-----------设置设计模式-------------
Dim sDocPath As String
sDocPath = oActDoc.Path '获取当前文档保存路径

DicPAll.RemoveAll

GetAllPN oActDoc.Product

Dim Num As Long

'Num = Int(txtStartNum.Text)
Num = CLng(txtStartNum.Text)
On Error Resume Next
Set oSel = Sel("Product")
If Err.Number <> 0 Then
Exit Sub
End If
Err.Clear

Dim oMyProducts
Set oMyProducts = oSel.Products
GetTargetDic oSel


If oMyProducts.Count <> 0 Then
    CATIA.DisplayFileAlerts = False

    Dim i As Integer
    Dim oProduct9 As Object

    For Each oProduct9 In oMyProducts                           '##0
        Dim NewPNi As String
          NewPNi = txtPreFix.Text & PreZero(CStr(Num), Len(txtStartNum.Text)) & txtSufFix.Text '第一种批量修改的情况
            Do Until Not DicPAll.exists(NewPNi)   '如果新料号已经存在于当前内存，则跳过取下一个
                    If DicP.exists(NewPNi) Then   '与新命名同规则的不需要改名
                    DicP.Remove NewPNi
                    End If
                Num = Num + Int(txtStep.Text)
                NewPNi = txtPreFix.Text & PreZero(CStr(Num), Len(txtStartNum.Text)) & txtSufFix.Text '第一种批量修改的情况
            Loop
        If Not ProductIsComponent(oProduct9) Then                '--1
                       Err.Clear
                       If DicP.exists(oProduct9.PartNumber) Then
                           If Err.Number = 0 Then '已经加载才修改
               
                               Dim NewName As String

                                    NewName = Replace(oProduct9.Name, oProduct9.PartNumber, NewPNi) '如果实例名称包含料号，则将实例名称中料号替换为新
                        
                                    oProduct9.PartNumber = NewPNi
                                    'MsgBox oProduct9.Name
                                    oProduct9.Name = NewName
                                    Err.Clear
                                    On Error Resume Next
                                    If chkOldFolder.Value = True Then
                                    sDocPath = oProduct9.ReferenceProduct.Parent.Path
'                                    MsgBox sDocPath
                                    End If
                                    oProduct9.ReferenceProduct.Parent.SaveAs sDocPath & "\" & NewPNi
                                    If Err.Number <> 0 Then
                                        Dim NewPN2 As String
                                        Randomize
                                        NewPN2 = NewPNi & "_" & Int(100 * Rnd)
                                        MsgBox sDocPath & "\" & NewPNi & " 无法保存,可能已经存在同名文件" & vbCrLf & _
                                               "将保存为" & sDocPath & "\" & NewPN2
                                        oProduct9.PartNumber = NewPN2
                                        oProduct9.ReferenceProduct.Parent.SaveAs sDocPath & "\" & NewPN2
                                    End If
                                    DicPAll.Add oProduct9.PartNumber, oProduct9.ReferenceProduct.Parent.FullName
                                    Err.Clear
                                 Num = Num + Int(txtStep.Text)
                           End If
                         End If
        End If                                            '--1
        
    Next                                               '##0
    
End If


'记录需要删除的旧文件
If cbxDelOld.Value = True Then
    Dim AArKill
    AArKill = DicP.keys()

    On Error Resume Next
    '删除旧文件
    
    If cbxDelOld.Value = True And UBound(AArKill) <> -1 Then
        If MsgBox("确实要删除旧文件吗? 删除后无法找回，请谨慎操作", vbYesNo + vbInformation, "零件批量重命名") = vbYes Then
        
            Dim kk As Integer
            For kk = 0 To UBound(AArKill)
            Kill DicP.Item(AArKill(kk))
            Next
        End If
    End If
End If

DicP.RemoveAll
Err.Clear
'添加空白零件
If chkBlankPart.Value = True Then

MsgBox "将新增 " & Int(Val(txtBlankQty.Text)) & "个空白零件" & vbCrLf & _
        "编号从 " & txtPreFix.Text & PreZero(CStr(Num), Len(txtStartNum.Text)) & txtSufFix.Text & "开始", vbOKOnly, "臭豆腐工具箱CATIA版"
Dim h As Integer
For h = 1 To Int(Val(txtBlankQty.Text))
 Dim NewPN3 As String
 Dim NewPart As Object
 NewPN3 = txtPreFix.Text & PreZero(CStr(Num), Len(txtStartNum.Text)) & txtSufFix.Text
 Set NewPart = oMyProducts.AddNewComponent("Part", NewPN3)
 NewPart.ReferenceProduct.Parent.SaveAs oSel.ReferenceProduct.Parent.Path & "\" & NewPN3 & ".CATPart"
 Num = Num + Int(txtStep.Text)
 Set NewPart = Nothing
Next
End If
chkBlankPart.Value = False
txtStartNum.Text = PreZero(CStr(Num), Len(txtStartNum.Text))
oActDoc.Save
If Err.Number <> 0 Then
MsgBox "保存时出了一点问题,可能的原因: " & vbCrLf & "- 是否试图向零件(Part)中添加零件(Part),而不是向产品(Product)中添加零件(Part)?" & vbCrLf & "- 新建的文件名已经在系统中存在? " & vbCrLf & vbCrLf & "请Ctrl + S 保存!", vbOKOnly, "臭豆腐工具箱CATIA版"
End If
CATIA.DisplayFileAlerts = True

End Sub

Private Sub cmdReName2_Click()
'检查输入有效性
On Error Resume Next
'*****************运行环境检查*********************
If txtFind = "" Then
MsgBox lblFind.Caption & "不能为空"
Exit Sub
End If

If IsNumeric(txtStart) = False Then
MsgBox "填写错误"
Exit Sub
End If
If CLng(txtStart) < 1 Then
MsgBox "至少要从第1个字符开始"
Exit Sub
End If

If IsNumeric(txtCount) = False Or CLng(txtCount) = 0 Then
MsgBox "填写错误"
Exit Sub
End If
'*****************运行环境检查*********************

Dim casesens
If chkCASE.Value = True Then
casesens = 1
Else
casesens = 0
End If

IntCATIA
'-----------设置设计模式-------------
CATIA.DisplayFileAlerts = False
'Dim settingControllers1
'Dim cacheSettingAtt1
'Set settingControllers1 = CATIA.SettingControllers
'Set cacheSettingAtt1 = settingControllers1.Item("CATSysCacheSettingCtrl")
'cacheSettingAtt1.ActivationMode = 1
'Set cacheSettingAtt1 = Nothing
'Set settingControllers1 = Nothing
Dim oProdxx
Set oProdxx = oActDoc.Product
oProdxx.ApplyWorkMode DESIGN_MODE
Set oProdxx = Nothing
'-----------设置设计模式-------------
Dim sDocPath As String
sDocPath = oActDoc.Path '获取当前文档保存路径

Err.Clear
DicPAll.RemoveAll
GetAllPN oActDoc.Product

Set oSel = Sel("Product")
If Err.Number <> 0 Then
Exit Sub
End If
Err.Clear

Dim oMyProducts
Set oMyProducts = oSel.Products
GetTargetDic oSel


If oMyProducts.Count <> 0 Then
    CATIA.DisplayFileAlerts = False

    Dim i As Integer

    For i = 1 To oMyProducts.Count
        Dim NewPNi As String
        NewPNi = Replace(oMyProducts.Item(i).PartNumber, RTrim(txtFind), RTrim(txtReplaceWith), CLng(txtStart), CInt(txtCount), casesens)
        If DicPAll.exists(NewPNi) Then
            DicP.Remove NewPNi
        End If
    Err.Clear
    Next

    For i = 1 To oMyProducts.Count
      If Not ProductIsComponent(oMyProducts.Item(i)) Then '--1
      
       If DicP.exists(oMyProducts.Item(i).PartNumber) Then
                If Err.Number = 0 Then
                   Dim NewName As String
                   Dim NewPN As String
                   NewPN = Replace(oMyProducts.Item(i).PartNumber, RTrim(txtFind), RTrim(txtReplaceWith), CLng(txtStart), CInt(txtCount), casesens)
                   NewName = Replace(oMyProducts.Item(i).Name, oMyProducts.Item(i).PartNumber, NewPN)
                   
                   oMyProducts.Item(i).PartNumber = NewPN
                   oMyProducts.Item(i).Name = NewName
                   If chkOldFolder.Value = True Then
                    sDocPath = oMyProducts.Item(i).ReferenceProduct.Parent.Path
'                                    MsgBox sDocPath
                    End If
                   
                   Err.Clear
                       If Not DicPAll.exists(NewPN) Then
                       SaveNew oMyProducts.Item(i), NewPN, NewName, sDocPath
                       End If
                 End If
            Err.Clear
        End If
       End If                                             '--1
    Next
End If


'记录需要删除的旧文件
If cbxDelOld.Value = True Then
        Dim AArKill
        AArKill = DicP.keys()
    
    
    
    '删除旧文件
'    MsgBox cbxDelOld.Value
'    MsgBox UBound(AArKill)
    
    If cbxDelOld.Value = True And UBound(AArKill) <> -1 Then
        If MsgBox("确实要删除旧文件吗? 删除后无法找回，请谨慎操作", vbYesNo + vbInformation, "零件批量重命名") = vbYes Then
        
            Dim kk As Integer
            For kk = 0 To UBound(AArKill)
            Kill DicP.Item(AArKill(kk))
            Next
        End If
    End If
End If
DicP.RemoveAll
oActDoc.Save
If Err.Number <> 0 Then
MsgBox "请Ctrl + S 保存!", vbOKOnly, "臭豆腐工具箱CATIA版"
End If
CATIA.DisplayFileAlerts = True

End Sub

Private Sub cmdPick3_Click()
cbxIgnoreHide.Enabled = False
'检查输入有效性
On Error Resume Next
'*****************运行环境检查*********************
If txtFind = "" Then
MsgBox lblFind.Caption & "不能为空"
Exit Sub
End If

If IsNumeric(txtStart) = False Then
MsgBox "填写错误"
Exit Sub
End If
If CLng(txtStart) < 1 Then
MsgBox "至少要从第1个字符开始"
Exit Sub
End If

If IsNumeric(txtCount) = False Or CInt(txtCount) = 0 Then
MsgBox "填写错误"
Exit Sub
End If
'*****************运行环境检查*********************

Dim casesens
If chkCASE.Value = True Then
casesens = 1
Else
casesens = 0
End If

IntCATIA
'-----------设置设计模式-------------
CATIA.DisplayFileAlerts = False
'Dim settingControllers1
'Dim cacheSettingAtt1
'Set settingControllers1 = CATIA.SettingControllers
'Set cacheSettingAtt1 = settingControllers1.Item("CATSysCacheSettingCtrl")
'cacheSettingAtt1.ActivationMode = 1
'Set cacheSettingAtt1 = Nothing
'Set settingControllers1 = Nothing
Dim oProdxx
Set oProdxx = oActDoc.Product
oProdxx.ApplyWorkMode DESIGN_MODE
Set oProdxx = Nothing
'-----------设置设计模式-------------
Dim sDocPath As String
sDocPath = oActDoc.Path '获取当前文档保存路径

On Error Resume Next
DicPAll.RemoveAll
GetAllPN oActDoc.Product

Dim s2, InputObjectType(), Status, FlagFinish, i
Set s2 = CATIA.ActiveDocument.Selection
s2.Clear
DicP.RemoveAll
FlagFinish = 0

        ReDim InputObjectType(0)
        InputObjectType(0) = "Product"
Do Until FlagFinish = 1

        Status = s2.SelectElement3(InputObjectType, "Select the " & InputObjectType(0) & ". Press Esc to finish the selection.", 0, 1, 0)

        If (Status = "Cancel") Then
            FlagFinish = 1
            If MsgBox("将退出点选修改模式，要保存修改么?", vbYesNo, "零件批量重命名") = vbYes Then
            oActDoc.Save
            s2.Clear
            cbxIgnoreHide.Enabled = True
            End If
        End If
        
        For i = 1 To s2.Count2
            If Not ProductIsComponent(s2.Item2(i).Value) Then
               
               '记录修改的文件，准备删除
                If Not DicP.exists(s2.Item2(i).Value.PartNumber) Then
                    If Err.Number = 0 Then
                       
                        'MsgBox S2.Item2(i).Value.PartNumber & vbCrLf & S2.Item2(i).Value.ReferenceProduct.Parent.FullName
                        DicP.Add s2.Item2(i).Value.PartNumber, s2.Item2(i).Value.ReferenceProduct.Parent.FullName
                        Dim NewPN As String
                        Dim NewName As String
                        NewPN = Replace(s2.Item2(i).Value.PartNumber, RTrim(txtFind), RTrim(txtReplaceWith), CLng(txtStart), CInt(txtCount), casesens)
                        NewName = Replace(s2.Item2(i).Value.Name, s2.Item2(i).Value.PartNumber, NewPN)
                        If chkOldFolder.Value = True Then
                            sDocPath = s2.Item2(i).Value.ReferenceProduct.Parent.Path
'                                    MsgBox sDocPath
                        End If
                     
                     
                        If Not DicPAll.exists(NewPN) Then
                        SaveNew s2.Item2(i).Value, NewPN, NewName, sDocPath
                        End If
                    End If
                End If
                Err.Clear
            End If
        Next
    s2.Clear

Loop

'记录需要删除的旧文件
If cbxDelOld.Value = True Then
        Dim AArKill
        AArKill = DicP.keys()
    
    If cbxDelOld.Value = True And UBound(AArKill) <> -1 Then
        If MsgBox("确实要删除旧文件吗? 删除后无法找回，请谨慎操作", vbYesNo + vbInformation, "零件批量重命名") = vbYes Then
        
            Dim kk As Integer
            For kk = 0 To UBound(AArKill)
            'MsgBox AArKill(kk)
            Kill DicP.Item(AArKill(kk))
            Next
        End If
    End If
End If
DicP.RemoveAll
FlagFinish = 0
cbxIgnoreHide.Enabled = True
Set s2 = Nothing


End Sub

Private Sub CurrFolder_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
Open_Current_Folder.CATMain
End Sub


Function Hidden(oProduct9 As Product) As Boolean
Hidden = False

Dim Sel9 As Object

Set Sel9 = oActDoc.Selection
Sel9.Clear

Dim VisPropertyset As Object
Dim showstate
Sel9.Add oProduct9
Set VisPropertyset = Sel9.VisProperties
VisPropertyset.GetShow showstate
If showstate = 1 Then    '1为隐藏
Hidden = True
End If

Sel9.Clear
Set Sel9 = Nothing
Set oProduct9 = Nothing
Set VisPropertyset = Nothing

End Function

Sub GetTargetDic(oProd8 As Product)

If oProd8.Products.Count = 0 Then
MsgBox "所选没有子零件"
Exit Sub
End If

DicP.RemoveAll

Dim i As Integer
Dim FailLoad As Integer

For i = 1 To oProd8.Products.Count
On Error Resume Next

If Not ProductIsComponent(oProd8.Products.Item(i)) Then

   If Not DicP.exists(oProd8.Products.Item(i).PartNumber) Then    ' 统计总共需修改数量
      If cbxIgnoreHide.Value = False Then  '当忽略隐藏零件未勾选,不管零件是否隐藏都添加
         DicP.Add oProd8.Products.Item(i).PartNumber, oProd8.Products.Item(i).ReferenceProduct.Parent.FullName
      Else
         If Not Hidden(oProd8.Products.Item(i)) Then              '当忽略隐藏零件勾选，且零件未隐藏时添加
            DicP.Add oProd8.Products.Item(i).PartNumber, oProd8.Products.Item(i).ReferenceProduct.Parent.FullName
         End If
      End If
       
   End If
    If Err.Number <> 0 Then
        FailLoad = FailLoad + 1
        Err.Clear
    End If
End If

Next

If FailLoad <> 0 Then
MsgBox "共有 " & FailLoad & "个无法重命名,文件可能未正确加载!"
End If

End Sub
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

End Function





Function PreZero(s As String, di As Integer)
If di > 1 And di < 20 Then
    Do While Len(s) < CInt(di)
         s = "0" & CStr(s)
    Loop
End If
PreZero = s
End Function

Private Sub UserForm_Initialize()

Set DicP = CreateObject("Scripting.Dictionary") '需要操作改名的字典
Set DicPAll = CreateObject("Scripting.Dictionary") '当前活动文件的所有料号字典

IntCATIA
'-----------设置设计模式-------------
CATIA.DisplayFileAlerts = False
'Dim settingControllers1
'Dim cacheSettingAtt1
'Set settingControllers1 = CATIA.SettingControllers
'Set cacheSettingAtt1 = settingControllers1.Item("CATSysCacheSettingCtrl")
'cacheSettingAtt1.ActivationMode = 1
'Set cacheSettingAtt1 = Nothing
'Set settingControllers1 = Nothing
Dim oProdxx
Set oProdxx = oActDoc.Product
oProdxx.ApplyWorkMode DESIGN_MODE
Set oProdxx = Nothing
'-----------设置设计模式-------------
If TypeName(oActDoc) <> "ProductDocument" Then
        MsgBox "当前文件不是CATIA产品 !" & vbCrLf & "此命令只能在产品文件中运行!"
        Exit Sub

End If

GetAllPN oActDoc.Product
If UnloadQty <> 0 Then
   MsgBox "有" & UnloadQty & "个文件没有正确加载,程序可继续，但请关注未加载的文件！"
End If

ReadConf
End Sub
Private Sub txtBlankQty_KeyPress(ByVal KeyANSI As MSForms.ReturnInteger)
    Select Case KeyANSI
          Case Asc("0") To Asc("9")
          'Case Asc("-")

          Case Else
            MsgBox "只能输入正整数！"
    End Select
End Sub
Private Sub txtCount_KeyPress(ByVal KeyANSI As MSForms.ReturnInteger)
    Select Case KeyANSI
          Case Asc("0") To Asc("9")
          'Case Asc("-")

          Case Else
            MsgBox "只能输入正整数！"
            
    End Select
End Sub
Private Sub txtStart_KeyPress(ByVal KeyANSI As MSForms.ReturnInteger)
    Select Case KeyANSI
          Case Asc("0") To Asc("9")
          'Case Asc("-")

          Case Else
            MsgBox "只能输入正整数！"
            
    End Select
End Sub
Private Sub txtStartNum_KeyPress(ByVal KeyANSI As MSForms.ReturnInteger)
    Select Case KeyANSI
          Case Asc("0") To Asc("9")
          'Case Asc("-")

          Case Else
            MsgBox "只能输入数字！"
    End Select
End Sub
Private Sub txtStep_KeyPress(ByVal KeyANSI As MSForms.ReturnInteger)
    Select Case KeyANSI
          Case Asc("0") To Asc("9")
          'Case Asc("-")

          Case Else
            MsgBox "只能输入数字！"
    End Select
End Sub

Function Sel(objType As String, Optional ObjType2 As String, Optional ObjType3 As String)
Dim s2, InputObjectType(), Status, NotSel, obj, Picked
Set s2 = CATIA.ActiveDocument.Selection
s2.Clear

If ObjType2 = "" Then
        ReDim InputObjectType(0)
        InputObjectType(0) = objType
        ElseIf ObjType3 = "" Then
                ReDim InputObjectType(1)
                InputObjectType(0) = objType
                InputObjectType(1) = ObjType2
                Else
                    ReDim InputObjectType(2)
                    InputObjectType(0) = objType
                    InputObjectType(1) = ObjType2
                    InputObjectType(2) = ObjType3
End If

Dim indication, n
indication = ""
For n = 0 To UBound(InputObjectType)
    indication = indication & "," & InputObjectType(n)
Next

Picked = False
NotSel = True
    Do While NotSel
        Status = s2.SelectElement2(InputObjectType, "Select the " & indication, False)
        If (Status = "Cancel") Then
            Exit Function
        ElseIf (Status = "Redo" And Not Picked) Then
               ElseIf (Status = "Undo") Then
                        Exit Function
                    ElseIf (Status <> "Redo") Then Set obj = s2.Item(1).Value
                              Picked = True
                              NotSel = False
        End If
    Loop
    s2.Clear
    Set Sel = obj
End Function
Sub GetPN(oProduct) ' 输出当前Part或Product料号字典
On Error Resume Next
If DicPAll.exists(oProduct.ReferenceProduct.Parent.Name) Then ' 测试是否加载这个文件
        If Err.Number <> 0 Then
        UnloadQty = UnloadQty + 1
        'MsgBox "有文件未加载，将跳过属性取值！"
        End If
    On Error GoTo 0
    Exit Sub
End If

If Not DicPAll.exists(oProduct.PartNumber) Then
DicPAll.Add oProduct.PartNumber, oProduct.ReferenceProduct.Parent.FullName
End If
End Sub
Sub GetAllPN(oProduct)
GetPN oProduct
Dim i
If oProduct.Products.Count <> 0 Then
    For i = 1 To oProduct.Products.Count
       GetPN oProduct.Products.Item(i)
    Next
End If

End Sub
Sub SaveNew(oProduct99, sPN, sName, sFolder)
On Error Resume Next
    oProduct99.PartNumber = sPN
    'MsgBox oProduct9.Name
    oProduct99.Name = sName
    Err.Clear
    On Error Resume Next
    oProduct99.ReferenceProduct.Parent.SaveAs sFolder & "\" & sPN
    If Err.Number <> 0 Then
        Dim NewPN2 As String
        Randomize
        NewPN2 = sPN & "_" & Int(100 * Rnd)
        MsgBox sFolder & "\" & sPN & " 无法保存,可能已经存在同名文件" & vbCrLf & _
               "将保存为" & sFolder & "\" & NewPN2
        oProduct99.PartNumber = NewPN2
        oProduct99.ReferenceProduct.Parent.SaveAs sFolder & "\" & NewPN2
    End If
    DicPAll.Add oProduct99.PartNumber, oProduct99.ReferenceProduct.Parent.FullName

End Sub
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
    read_OK = GetPrivateProfileString("零件批量重命名", "忽略隐藏零件", "True", read2, 256, sConfpath & "\Conf1.ini")
    cbxIgnoreHide.Value = read2
    read2 = String(255, 0)
    read_OK = GetPrivateProfileString("零件批量重命名", "删除旧文件", "False", read2, 256, sConfpath & "\Conf1.ini")
    cbxDelOld.Value = read2
    read2 = String(255, 0)
    read_OK = GetPrivateProfileString("零件批量重命名", "保存在原文件夹", "True", read2, 256, sConfpath & "\Conf1.ini")
    chkOldFolder.Value = read2
    read2 = String(255, 0)
    read_OK = GetPrivateProfileString("零件批量重命名", "名称前缀", "CDF-", read2, 256, sConfpath & "\Conf1.ini")
    txtPreFix.Text = read2
    read2 = String(255, 0)
    read_OK = GetPrivateProfileString("零件批量重命名", "名称后缀", "", read2, 256, sConfpath & "\Conf1.ini")
    txtSufFix.Text = read2
    read2 = String(255, 0)
    read_OK = GetPrivateProfileString("零件批量重命名", "增加前缀", "", read2, 256, sConfpath & "\Conf1.ini")
    txtAddPre.Text = read2
    read2 = String(255, 0)
    read_OK = GetPrivateProfileString("零件批量重命名", "增加后缀", "", read2, 256, sConfpath & "\Conf1.ini")
    txtAddSuf.Text = read2
    read2 = String(255, 0)
    read_OK = GetPrivateProfileString("零件批量重命名", "查找字符串", "CDF-", read2, 256, sConfpath & "\Conf1.ini")
    txtFind.Text = read2
    read2 = String(255, 0)
    read_OK = GetPrivateProfileString("零件批量重命名", "替换为", "", read2, 256, sConfpath & "\Conf1.ini")
    txtReplaceWith.Text = read2
    read2 = String(255, 0)
    read_OK = GetPrivateProfileString("零件批量重命名", "不区分大小写", "True", read2, 256, sConfpath & "\Conf1.ini")
    chkCASE.Value = read2

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
write1 = WritePrivateProfileString("零件批量重命名", "忽略隐藏零件", CStr(cbxIgnoreHide.Value), sConfpath & "\Conf1.ini")
write1 = WritePrivateProfileString("零件批量重命名", "删除旧文件", CStr(cbxDelOld.Value), sConfpath & "\Conf1.ini")
write1 = WritePrivateProfileString("零件批量重命名", "保存在原文件夹", CStr(chkOldFolder.Value), sConfpath & "\Conf1.ini")
write1 = WritePrivateProfileString("零件批量重命名", "名称前缀", txtPreFix.Text, sConfpath & "\Conf1.ini")
write1 = WritePrivateProfileString("零件批量重命名", "名称后缀", txtSufFix.Text, sConfpath & "\Conf1.ini")
write1 = WritePrivateProfileString("零件批量重命名", "增加前缀", txtAddPre.Text, sConfpath & "\Conf1.ini")
write1 = WritePrivateProfileString("零件批量重命名", "增加后缀", txtAddSuf.Text, sConfpath & "\Conf1.ini")
write1 = WritePrivateProfileString("零件批量重命名", "查找字符串", txtFind.Text, sConfpath & "\Conf1.ini")
write1 = WritePrivateProfileString("零件批量重命名", "替换为", txtReplaceWith.Text, sConfpath & "\Conf1.ini")
write1 = WritePrivateProfileString("零件批量重命名", "不区分大小写", CStr(chkCASE.Value), sConfpath & "\Conf1.ini")
End Sub

Private Sub UserForm_Terminate()
SaveConf
End Sub

Private Sub WidthPlus_Click()
Dim addWidth
If (Batch_ReName.Width > 300) Then
       addWidth = -200
  ElseIf (Batch_ReName.Width < 200) Then
       addWidth = 200
End If
Batch_ReName.Left = Batch_ReName.Left - addWidth
Batch_ReName.Width = Batch_ReName.Width + addWidth
WidthPlus.Width = WidthPlus.Width + addWidth / 2
WidthPlus.Left = WidthPlus.Left + addWidth / 2

frmReName.Width = frmReName.Width + addWidth
txtPreFix.Width = txtPreFix.Width + addWidth
txtStartNum.Width = txtStartNum.Width + addWidth
txtStep.Width = txtStep.Width + addWidth
txtSufFix.Width = txtSufFix.Width + addWidth
txtBlankQty.Width = txtBlankQty.Width + addWidth
cmdReName.Width = cmdReName.Width + addWidth / 2
cmdPick1.Width = cmdPick1.Width + addWidth / 2
cmdPick1.Left = cmdPick1.Left + addWidth / 2

frmAddPreSuf.Width = frmAddPreSuf.Width + addWidth
txtAddPre.Width = txtAddPre.Width + addWidth
txtAddSuf.Width = txtAddSuf.Width + addWidth
cmdAddPreSuf.Width = cmdAddPreSuf.Width + addWidth / 2
cmdPick2.Width = cmdPick2.Width + addWidth / 2
cmdPick2.Left = cmdPick2.Left + addWidth / 2

frmReplace.Width = frmReplace.Width + addWidth
txtFind.Width = txtFind.Width + addWidth
txtReplaceWith.Width = txtReplaceWith.Width + addWidth
txtStart.Width = txtStart.Width + addWidth
txtCount.Width = txtCount.Width + addWidth
cmdReName2.Width = cmdReName2.Width + addWidth / 2
cmdPick3.Width = cmdPick3.Width + addWidth / 2
cmdPick3.Left = cmdPick3.Left + addWidth / 2

cmdInsName.Width = cmdInsName.Width + addWidth
cmdInsName2.Width = cmdInsName2.Width + addWidth / 3
cmdInsName3.Width = cmdInsName3.Width + addWidth / 3
cmdInsName3.Left = cmdInsName3.Left + addWidth / 3
cmdInsName4.Width = cmdInsName4.Width + addWidth / 3
cmdInsName4.Left = cmdInsName4.Left + 2 * addWidth / 3

End Sub
