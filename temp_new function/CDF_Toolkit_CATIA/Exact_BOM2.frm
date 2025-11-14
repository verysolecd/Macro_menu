VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Exact_BOM2 
   Caption         =   "导出BOM2"
   ClientHeight    =   9240.001
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   3315
   OleObjectBlob   =   "Exact_BOM2.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "Exact_BOM2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim OutPutFolder As String
Dim intLevelGlobal As Integer
Dim intIndexGlobal As Integer
Dim intIndexPL As Integer
Dim ExcelForWrite As Object
Dim BOM2()  '存储BOM2的二维数组
Dim PL2()     '存储PL2的二维数组
Dim UserProperties '存储用户属性名称
Dim pDic As Object  '属性字典,key为料号，值为一维数组，上标为选择的个数-1（不包含永远被选中的前4项
Dim FrontViewDirection(5) As Double
Dim oDicPL As Object  '属性字典,key为料号，值为一维数组，永远被选中的前4项
Dim iImgCol As Integer '图片列,值为1或者0

#If VBA7 Then
Private Declare PtrSafe Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As LongPtr
Private Declare PtrSafe Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As LongPtr
#Else
Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Private Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long
#End If





'***************************执行导出BOM2******************************
Private Sub cmdRun1_Click()


OutPutFolder = oCATVBA_Folder("Temp").Path


'-----------使用高速缓存------------
On Error Resume Next
IntCATIA
'MsgBox "IntCATIA OK"
CATIA.DisplayFileAlerts = False
'CATIA.Interactive = False
'Dim settingControllers1
'Dim cacheSettingAtt1
'Set settingControllers1 = CATIA.SettingControllers
'Set cacheSettingAtt1 = settingControllers1.Item("CATSysCacheSettingCtrl")
'cacheSettingAtt1.ActivationMode = 1
'Set cacheSettingAtt1 = Nothing
'Set settingControllers1 = Nothing
If Err.Number <> 0 Then
'Unload Exact_BOM2
MsgBox "无法与CATIA建立连接!" & vbCrLf & "你是否同时安装了不同版本的CATIA？"
Exit Sub
End If
If TypeName(oActDoc) <> "ProductDocument" Then
'    Unload Exact_BOM2
    MsgBox "当前文件不是CATIA产品 !" & vbCrLf & "此命令只能在装配工作台中运行!"
    lstPro.Clear
    Exit Sub
End If
'MsgBox TypeName(oActDoc)
If lstPro.ListCount = 0 Then
    Unload Exact_BOM2
Exact_BOM2.Show vbModeless
Exact_BOM2.Left = Exact_BOM2.Left * 2
'    MsgBox "请重新运行命令以便选择要输出的属性"
    Exit Sub

End If

Dim arrProX
arrProX = Array()
Set pDic = CreateObject("Scripting.Dictionary")  '属性字典,key为料号，值为一维数组，上标为选择的个数-1（不包含永远被选中的前4项
ReDim BOM2(1000, 3)
ReDim PL2(1000, 3)
BOM2(0, 0) = lstPro.List(0)
BOM2(0, 1) = lstPro.List(1)
BOM2(0, 2) = lstPro.List(2)
BOM2(0, 3) = lstPro.List(3)
PL2(0, 0) = lstPro.List(0)
PL2(0, 1) = "多件质量 Mass(xN)(g)"
PL2(0, 2) = lstPro.List(2)
PL2(0, 3) = lstPro.List(3)


Dim n As Integer
For n = 4 To lstPro.ListCount - 1
    '如果被选择
    If lstPro.Selected(n) = True Then
    '则加入arrProX数组
    ReDim Preserve arrProX(UBound(arrProX) + 1)
            arrProX(UBound(arrProX)) = lstPro.List(n)
    ReDim Preserve BOM2(1000, UBound(BOM2, 2) + 1)
            BOM2(0, UBound(BOM2, 2)) = lstPro.List(n)
    ReDim Preserve PL2(1000, UBound(PL2, 2) + 1)
            PL2(0, UBound(PL2, 2)) = lstPro.List(n)
    End If
Next
Dim sc As String
sc = cmdRun1.Caption
cmdRun1.Caption = "请稍候..."
cmdRun1.BackColor = &HFF&
ExactBOMPL
Set pDic = Nothing
cmdRun1.Caption = sc
cmdRun1.BackColor = &H8000000F
CATIA.Interactive = True
CATIA.RefreshDisplay = True
CATIA.StatusBar = "导出BOM完成！"
End Sub



'***************************主窗口初始化******************************
Private Sub UserForm_Initialize()

IntCATIA

If TypeName(oActDoc) <> "ProductDocument" Then
        MsgBox "此命令只能在装配工作台中运行!" & vbCrLf & "请打开产品数模后重新运行此命令"
        Exit Sub
        Unload Me
End If
lstPro.AddItem "序号 Index"
lstPro.AddItem "层级 Level"
lstPro.AddItem "零件编号 Part Number"
lstPro.AddItem "数量 Quantities"
lstPro.AddItem "实例名称 Instance Name"
lstPro.AddItem "实例描述 Component Description"
lstPro.AddItem "零件名称 Definition"
lstPro.AddItem "零件描述 Part Description"
lstPro.AddItem "术语 Nomenclature"
lstPro.AddItem "版本 Rev"
lstPro.AddItem "来源 Source"
lstPro.AddItem "密度 Density(g/cm3)"
lstPro.AddItem "单件表面积 WetArea(cm2)"
lstPro.AddItem "单件体积 Volume(cm3)"
lstPro.AddItem "单件质量 Mass(gram)"
lstPro.AddItem "外形尺寸 Bounding Box(mm)"


UserProperties = Array()
GetUserProperties2 oActDoc.Product
If UBound(UserProperties) <> -1 Then
'MsgBox UBound(UserProperties)
Dim b As Integer
For b = 0 To UBound(UserProperties)
    lstPro.AddItem UserProperties(b)
Next
End If

lstProAlwaysCheckedItems
iImgCol = 0

ReadConf
End Sub

'***********************提取BOM和PL************************
Sub ExactBOMPL()
On Error Resume Next
    intIndexGlobal = 0       'Index for BOM
    intLevelGlobal = 0       'Level for BOM
    intIndexPL = 0          'Index for PartList sheet
    Set oDicPL = CreateObject("Scripting.Dictionary")  '属性字典,key为料号，永远被选中的前4项
    If TypeName(oActDoc) <> "ProductDocument" Then
        MsgBox "当前文件不是CATIA产品 !" & vbCrLf & "此命令只能在装配工作台中运行!"
        Exit Sub
    End If

    If chkImg.Value = True Then
    cmdRun1.Caption = "正在截图..."
'    hideConstraints oActDoc.Product
'    HideDatums oActDoc.Product
    HideDatumsConstraints oActDoc, 1
    CaptureChildrenImage oActDoc.Product, OutPutFolder
    End If
        Set ExcelForWrite = CreateObject("Excel.Application")
            ExcelForWrite.workbooks.Add
'            ExcelForWrite.Interactive = 0
'            ExcelForWrite.DisplayAlerts = 0
            ExcelForWrite.RollZoom = 0
            'ExcelForWrite.Worksheets.Add
            ExcelForWrite.ActiveSheet.Name = Left(oActDoc.Product.PartNumber, 20) & "_BOM"
            'ExcelForWrite.ActiveSheet.Columns("B:B").HorizontalAlignment = -4131 'xlLeft
            ExcelForWrite.ActiveSheet.Columns("C:C").NumberFormatLocal = "@"
            ExcelForWrite.Worksheets.Add
            ExcelForWrite.ActiveSheet.Name = Left(oActDoc.Product.PartNumber, 20) & "_PartUsage"
            ExcelForWrite.ActiveSheet.Columns("C:C").NumberFormatLocal = "@"
        If Err.Number <> 0 Then
            MsgBox "不能创建Excel文件, 程序退出"
            Exit Sub
        End If

Set ExcelForWrite = ExcelForWrite.Worksheets(Left(oActDoc.Product.PartNumber, 20) & "_BOM")

Extract_BOM oActDoc.Product
cmdRun1.Caption = "正在生成Excel表..."
CATIA.StatusBar = "正在生成Excel表..."
Dim iRowshift, iColshift As Integer
iRowshift = 7   '从第几行开始
iColshift = 1       '从第几列开始
ExcelForWrite.cells(iRowshift, iColshift).Resize(UBound(BOM2) + 1, UBound(BOM2, 2) + 1) = BOM2

If chkImgCol.Value = True Then
    iImgCol = 1
    ExcelForWrite.Columns("E:E").Insert
    ExcelForWrite.cells(iRowshift, 5).Value = "图片 Picture"
    ExcelForWrite.cells(iRowshift, 5).ColumnWidth = 15
    
Else
    iImgCol = 0
End If

Dim iRowPN, iColPN As Integer
iColPN = 2              '料号在弟3列，偏移之前
For iRowPN = 1 To intIndexGlobal     '第一列是名称，所以不从0开始
        If Not IsEmpty(ExcelForWrite.cells(iRowPN + iRowshift, iColPN + iColshift).Value) Then
                If chkImg.Value = True Then
                CommentImg ExcelForWrite.cells(iRowPN + iRowshift, iColPN + iColshift), OutPutFolder & "\" & RebuildFileName(ExcelForWrite.cells(iRowPN + iRowshift, iColPN + iColshift).Value) & ".jpg"
                    If chkImgCol.Value = True Then
                        ExcelForWrite.cells(iRowPN + iRowshift, 5).RowHeight = 40
                        ColImg ExcelForWrite, ExcelForWrite.cells(iRowPN + iRowshift, 5), OutPutFolder & "\" & RebuildFileName(ExcelForWrite.cells(iRowPN + iRowshift, iColPN + iColshift).Value) & ".jpg"
                    End If
                End If
        End If
Next
TableHeadFormat "BOM", oActDoc.Product

'ExcelForWrite.Parent.Parent.Visible = True

Set ExcelForWrite = ExcelForWrite.Parent.Worksheets(Left(oActDoc.Product.PartNumber, 20) & "_PartUsage")
Extract_PartList oActDoc.Product
ExcelForWrite.cells(iRowshift, iColshift).Resize(UBound(PL2) + 1, UBound(PL2, 2) + 1) = PL2
If chkImgCol.Value = True Then
    iImgCol = 1
    ExcelForWrite.Columns("E:E").Insert
    ExcelForWrite.cells(iRowshift, 5).Value = "图片 Picture"
    ExcelForWrite.cells(iRowshift, 5).ColumnWidth = 15
Else
    iImgCol = 0
End If

For iRowPN = 1 To intIndexPL    '第一列是名称，所以不从0开始
        If Not IsEmpty(ExcelForWrite.cells(iRowPN + iRowshift, iColPN + iColshift).Value) Then
                If chkImg.Value = True Then
                CommentImg ExcelForWrite.cells(iRowPN + iRowshift, iColPN + iColshift), OutPutFolder & "\" & RebuildFileName(ExcelForWrite.cells(iRowPN + iRowshift, iColPN + iColshift).Value) & ".jpg"
                 If chkImgCol.Value = True Then
                        ExcelForWrite.cells(iRowPN + iRowshift, 5).RowHeight = 40
                        ColImg ExcelForWrite, ExcelForWrite.cells(iRowPN + iRowshift, 5), OutPutFolder & "\" & RebuildFileName(ExcelForWrite.cells(iRowPN + iRowshift, iColPN + iColshift).Value) & ".jpg"
                 End If
                End If
        End If
Next
TableHeadFormat "PL", oActDoc.Product
ExcelForWrite.cells(iRowshift + 1, 5 + iImgCol).Select
ExcelForWrite.Parent.Parent.ActiveWindow.FreezePanes = True


Set ExcelForWrite = ExcelForWrite.Parent.Worksheets(Left(oActDoc.Product.PartNumber, 20) & "_BOM")
ExcelForWrite.Activate
ExcelForWrite.Parent.Parent.Visible = True
ExcelForWrite.cells(iRowshift + 1, 5 + iImgCol).Select
ExcelForWrite.Parent.Parent.ActiveWindow.FreezePanes = True
'ExcelForWrite.Parent.Parent.DisplayAlerts = 1

'On Error Resume Next
Err.Clear
Dim xlName As String
Randomize
xlName = OutPutFolder & "\" & "BOM_PL_" & Replace(oActDoc.Name, ".", "_") & "_" & Int(100 * Rnd) & ".xls"
'MsgBox xlName & vbCrLf & TypeName(ExcelForWrite.Parent)
ExcelForWrite.Parent.SaveAs xlName
'MsgBox "操作完成!" & vbCrLf & _
'            "文件保存在" & xlName
Set ExcelForWrite = Nothing
Err.Clear
End Sub
'***********************为Excel单元格添加注释************************
Private Sub CommentImg(ByVal oCells As Object, ByVal ImgName As String)   'ImgName is full path with extension name(.jpg ect)
    On Error Resume Next
    With oCells
       .AddComment
       .Comment.Visible = True
       .Comment.Text ""
       .Comment.Shape.Fill.UserPicture ImgName
       .Comment.Shape.Width = 120
       .Comment.Shape.Height = 80
       .Comment.Visible = False
    End With
    Err.Clear
End Sub
Private Sub ColImg(ByVal oSheet As Object, ByVal oRange As Object, ByVal sfilename As String)
On Error Resume Next
Dim oShape As Object
Set oShape = oSheet.Shapes.AddPicture(sfilename, 0, -1, oRange.Left + 1, oRange.Top + 1, oRange.Width - 1, oRange.Height - 1)
oShape.Placement = 1
Set oShape = Nothing
End Sub

'***********************永远选择的输出项，用户不能更改************************
Private Sub lstPro_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
lstProAlwaysCheckedItems
End Sub
'***********************永远选择的输出项，用户不能更改************************
Private Sub lstProAlwaysCheckedItems()
On Error Resume Next
        If lstPro.List(0) = "序号 Index" Then
        lstPro.Selected(0) = True
        End If
        If lstPro.List(1) = "层级 Level" Then
        lstPro.Selected(1) = True
        End If
        If lstPro.List(2) = "零件编号 Part Number" Then
        lstPro.Selected(2) = True
        End If
        If lstPro.List(3) = "数量 Quantities" Then
        lstPro.Selected(3) = True
        End If
Err.Clear
End Sub
'***********************根据属性名称取得属性值，默认属性************************
Function PropertyValue(ByVal oProduct1 As Object, ByVal PropertyName As String)
PropertyValue = ""
If TypeName(oProduct1) <> "Product" Then
Exit Function
End If
On Error Resume Next
Select Case PropertyName
    Case "实例名称 Instance Name"
        PropertyValue = oProduct1.Name
    Case "实例描述 Component Description"
        PropertyValue = oProduct1.DescriptionInst
    Case "零件名称 Definition"
        PropertyValue = oProduct1.Definition
    Case "零件描述 Part Description"
        PropertyValue = oProduct1.DescriptionRef
    Case "术语 Nomenclature"
        PropertyValue = oProduct1.Nomenclature
    Case "版本 Rev"
        PropertyValue = oProduct1.Revision
    Case "来源 Source"
        PropertyValue = oProduct1.Source
        Select Case oProduct1.Source
            Case 0
                    PropertyValue = "未定义/Unknown"
            Case 1
                    PropertyValue = "自制/Made"
            Case 2
                    PropertyValue = "外购/Bought"
        End Select
    Case "密度 Density(g/cm3)"
        PropertyValue = oProduct1.GetTechnologicalObject("Inertia").Density / 1000        'default kg/m^3, transfer to g/cm3
    Case "单件表面积 WetArea(cm2)"
        PropertyValue = Round(oProduct1.Analyze.WetArea / 100, 2)                'cm^2
    Case "单件体积 Volume(cm3)"
        PropertyValue = Round(oProduct1.Analyze.Volume / 1000, 2)               'cm^3
    Case "单件质量 Mass(gram)"
        PropertyValue = Round(oProduct1.Analyze.Mass * 1000, 2)                 'grameter
    Case "外形尺寸 Bounding Box(mm)"
        PropertyValue = GetBoxSize(oProduct1.ReferenceProduct)
    Case Else
       Dim k As Integer
       Dim oUserPara As Object 'Parameters
                Set oUserPara = oProduct1.ReferenceProduct.UserRefProperties
       For k = 1 To oUserPara.Count
       If PropertyName = GetUserPropStr(oUserPara.Item(k).Name) Then
                PropertyValue = oUserPara.Item(k).ValueAsString
       End If
       Next
End Select
Err.Clear
End Function
'************************************取得用户属性名称****************************************
Function GetUserPropStr(ByVal s As String)
GetUserPropStr = Right(s, Len(s) - InStrRev(s, "\"))
End Function
'***************************是否是零件******************************
Function is_Leaf(ByVal oProduct1)
    is_Leaf = (oProduct1.Products.Count = 0)
End Function
'***************************BoundingBox******************************
Function GetBoxSize(ByVal oProduct11) 'As String
'CATIA.Visible = False
GetBoxSize = ""
On Error Resume Next
CATIA.RefreshDisplay = False
Dim documents11, drawingDocument11, drawingSheets11, drawingSheet11, drawingViews1, drawingView1
Dim drawingViewGenerativeLinks1, drawingViewGenerativeBehavior1, drawingView2, drawingViewGenerativeLinks2, drawingViewGenerativeBehavior2
Set documents11 = CATIA.Documents
Set drawingDocument11 = documents11.Add("Drawing")
'drawingDocument11.Standard = catISO
Set drawingSheets11 = drawingDocument11.Sheets
Set drawingSheet11 = drawingSheets11.Item(1)
'drawingSheet11.PaperSize = catPaperA0
drawingSheet11.Scale2 = 1
'drawingSheet11.Orientation = catPaperLandscape
Set drawingViews1 = drawingSheet11.Views
Set drawingView1 = drawingViews1.Add("AutomaticNaming")
drawingView1.x = 10.5
drawingView1.y = 10.5
drawingView1.Scale2 = 1
Set drawingViewGenerativeLinks1 = drawingView1.GenerativeLinks
Set drawingViewGenerativeBehavior1 = drawingView1.GenerativeBehavior
drawingViewGenerativeBehavior1.Document = oProduct11
            drawingViewGenerativeBehavior1.DefineFrontView 0, 1, 0, 0, 0, 1
            If oProduct11.Products.Count = 0 Then ' 如果是零件，优先用户第一个轴系
            FVDirection oProduct11.Parent.Part
            drawingViewGenerativeBehavior1.DefineFrontView FrontViewDirection(0), FrontViewDirection(1), FrontViewDirection(2), FrontViewDirection(3), FrontViewDirection(4), FrontViewDirection(5)
            End If

'drawingViewGenerativeBehavior1.DefineFrontView 1, 0, 0, 0, 1, 0
Set drawingView2 = drawingViews1.Add("AutomaticNaming")
drawingView2.x = 500
drawingView2.y = 10.5
drawingView2.Scale2 = 1
Set drawingViewGenerativeLinks2 = drawingView2.GenerativeLinks
drawingViewGenerativeLinks1.CopyLinksTo drawingViewGenerativeLinks2
Set drawingViewGenerativeBehavior2 = drawingView2.GenerativeBehavior
drawingViewGenerativeBehavior2.DefineProjectionView drawingViewGenerativeBehavior1, catLeftView
drawingViewGenerativeBehavior1.Update
drawingViewGenerativeBehavior2.Update
'drawingView1.Activate
Dim L, h, W
L = 0
h = 0
W = 0
Dim oXY(4) 'As Double
drawingView1.Size oXY
Dim Xmin, Xmax, Ymin, Ymax
Xmin = oXY(0)
Xmax = oXY(1)
Ymin = oXY(2)
Ymax = oXY(3)
L = Round(Abs(Xmax - Xmin), 1)
h = Round(Abs(Ymax - Ymin), 1)
drawingView2.Size oXY
Xmin = oXY(0)
Xmax = oXY(1)
Ymin = oXY(2)
Ymax = oXY(3)
W = Round(Abs(Xmax - Xmin), 1)
Dim tmp
If L < W Then
   tmp = L
   L = W
   W = tmp
End If
If W < h Then
   tmp = W
   W = h
   h = tmp
End If
If L < W Then
   tmp = L
   L = W
   W = tmp
End If
GetBoxSize = CStr(L) + "x" + CStr(W) + "x" + CStr(h)
drawingDocument11.Close
Set documents11 = Nothing
Set drawingDocument11 = Nothing
Set drawingSheets11 = Nothing
Set drawingSheet11 = Nothing
Set drawingViews1 = Nothing
Set drawingViewGenerativeLinks1 = Nothing
Set drawingViewGenerativeBehavior1 = Nothing
Set drawingView2 = Nothing
Err.Clear
End Function
'***************************是否为部件******************************
Function ProductIsComponent(ByVal iProduct) As Boolean
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

'***************************Excel表头******************************

Sub TableHeadFormat(ByVal BOMorPL, ByVal oProduct1)
On Error Resume Next
With ExcelForWrite.cells.Font
        .Name = "Microsoft YaHei UI"
        .Size = 11
        .Strikethrough = False
        .Superscript = False
        .Subscript = False
        .OutlineFont = False
        .Shadow = False
        .Underline = 1 'xlUnderlineStyleNone
        .ThemeColor = 2 'xlThemeColorLight1
        .TintAndShade = 0
        .ThemeFont = 1 'xlThemeFontNone
End With

With ExcelForWrite.cells
        .HorizontalAlignment = -4108 'xlCenter
        .VerticalAlignment = -4160 'xlTop
        .WrapText = True
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = 1 'xlContext
        .MergeCells = False
        '.RowHeight = 20
End With
If iImgCol = 0 Then
ExcelForWrite.cells.RowHeight = 20
End If


With ExcelForWrite.Rows("1:1")
        .WrapText = False
        .Font.Size = 22
        .RowHeight = 35
        .HorizontalAlignment = -4131 'xlLeft
End With

With ExcelForWrite.Rows("3:6")
        .WrapText = False
        .HorizontalAlignment = -4131 'xlLeft
End With
With ExcelForWrite.Rows("7:7")
        .WrapText = False
        .Columns.AutoFit
End With

ExcelForWrite.cells(1, 3).Value = "物料清单/BILL OF MATERIAL"
ExcelForWrite.cells(3, 1).Value = "产品编号/PN:"
ExcelForWrite.cells(3, 3).Value = oProduct1.PartNumber
ExcelForWrite.cells(4, 1).Value = "产品名称/Product Name:"
ExcelForWrite.cells(4, 3).Value = oProduct1.Definition
ExcelForWrite.cells(5, 1).Value = "版本/Revision:"
ExcelForWrite.cells(5, 3).Value = oProduct1.Revision
ExcelForWrite.cells(6, 1).Value = "质量/Mass(g):"
ExcelForWrite.cells(6, 3).Value = Round(oProduct1.Analyze.Mass * 1000, 2)
Dim oDataArea As Object
Dim iRowshift, iColshift, bd As Integer
iRowshift = 7
iColshift = 1
Set oDataArea = ExcelForWrite.cells(iRowshift, iColshift).Resize(1 + intIndexGlobal, UBound(BOM2, 2) + 1 + iImgCol)
    If BOMorPL = "PL" Then
       ExcelForWrite.cells(1, 3).Value = "零件用量/PARTS USAGE"
       ExcelForWrite.cells(iRowshift + 1 + intIndexPL, 1).Value = "【此表" & "于" & Now() & "由臭豆腐工具箱CATIA版自动生成】【微信公众号：UG遇上CATIA】"
       ExcelForWrite.cells(iRowshift + 1 + intIndexPL, 1).WrapText = False
       ExcelForWrite.cells(iRowshift + 1 + intIndexPL, 1).HorizontalAlignment = -4131 'xlLeft
       ExcelForWrite.cells(iRowshift, iColshift).Resize(1, UBound(PL2, 2) + 1 + iImgCol).Interior.color = 15773696
       ExcelForWrite.cells(iRowshift, iColshift).Resize(1, UBound(PL2, 2) + 1 + iImgCol).AutoFilter
'        ExcelForWrite.Cells(iRowshift + 2 + intIndexPL, 1).Value = "【臭豆腐工具箱CATIA版的[属性补全]命令可快速增删改数模属性，欢迎下载体验】"
'       ExcelForWrite.Cells(iRowshift + 2 + intIndexPL, 1).WrapText = False
'       ExcelForWrite.Cells(iRowshift + 2 + intIndexPL, 1).HorizontalAlignment = -4131 'xlLeft

       
       
       
       Set oDataArea = ExcelForWrite.cells(iRowshift, iColshift).Resize(1 + intIndexPL, UBound(PL2, 2) + 1 + iImgCol)
    Else
       ExcelForWrite.cells(iRowshift + 1 + intIndexGlobal, 1).Value = "【此表" & "于" & Now() & "由臭豆腐工具箱CATIA版自动生成】【微信公众号：UG遇上CATIA】"
       ExcelForWrite.cells(iRowshift + 1 + intIndexGlobal, 1).WrapText = False
       ExcelForWrite.cells(iRowshift + 1 + intIndexGlobal, 1).HorizontalAlignment = -4131 'xlLeft
       ExcelForWrite.cells(iRowshift, iColshift).Resize(1, UBound(BOM2, 2) + 1 + iImgCol).Interior.color = 49407
       '20221203v073增加筛选和冻结单元格
       ExcelForWrite.cells(iRowshift, iColshift).Resize(1, UBound(BOM2, 2) + 1 + iImgCol).AutoFilter

       ExcelForWrite.cells(iRowshift + 1, iColshift + 1).Resize(UBound(BOM2) + 1, 1).HorizontalAlignment = -4131 'xlLeft
'       ExcelForWrite.Cells(iRowshift + 2 + intIndexGlobal, 1).Value = "【臭豆腐工具箱CATIA版的[属性补全]命令可快速增删改数模属性，欢迎下载体验】"
'       ExcelForWrite.Cells(iRowshift + 2 + intIndexGlobal, 1).WrapText = False
'       ExcelForWrite.Cells(iRowshift + 2 + intIndexGlobal, 1).HorizontalAlignment = -4131 'xlLeft
       
       
    End If
For bd = 7 To 12
With oDataArea.Borders(bd)
        .LineStyle = 1 'xlContinuous
        .ColorIndex = 1 'xlAutomatic
        .TintAndShade = 0
        .Weight = 2 ' xlThin
End With

Next



Err.Clear
End Sub
'***************************提取BOM******************************
Sub Extract_BOM(ByVal oProduct1)

    Dim i As Integer
    Dim oProducts1
   On Error Resume Next
    Dim pDicX As Object  '属性字典,key为料号，值为一维数组，永远被选中的前4项
    Set pDicX = CreateObject("Scripting.Dictionary")  '属性字典,key为料号，永远被选中的前4项
    Set oProducts1 = oProduct1.Products
    intLevelGlobal = intLevelGlobal + 1
    If is_Leaf(oProduct1) Then
         Exit Sub
    End If
    For i = 1 To oProducts1.Count
    oProducts1.Item(i).ApplyWorkMode 2
'    MsgBox oProducts1.Item(i).Name
'    MsgBox oProducts1.Item(i).PartNumber
           Err.Clear
           Dim Unload1, PN
           PN = oProducts1.Item(i).PartNumber
           If Err.Number <> 0 Then
           Unload1 = True
           PN = "未加载无法取得数据"
           Else
           Unload1 = False
           End If
           Err.Clear

           If Unload1 = False Then
           cmdRun1.Caption = "操作" & oProducts1.Item(i).PartNumber & "..."
           CATIA.StatusBar = "操作" & oProducts1.Item(i).PartNumber & "..."
           If i Mod 2 = 0 Then
           cmdRun1.BackColor = &H8000000D '&H8000000F
           Else
           cmdRun1.BackColor = &HFFFF&
           End If
                     If Not pDicX.exists(oProducts1.Item(i).PartNumber) Then          '----------------1
                       Dim ProExist
                       intIndexGlobal = intIndexGlobal + 1
                       BOM2(intIndexGlobal, 0) = intIndexGlobal 'Index
                       BOM2(intIndexGlobal, 1) = String(intLevelGlobal - 1, "○") & "●" & intLevelGlobal                   'Level
                       BOM2(intIndexGlobal, 2) = oProducts1.Item(i).PartNumber               'Same as dictory key
                       BOM2(intIndexGlobal, 3) = 1                                                                         'Qty
                       pDicX.Add BOM2(intIndexGlobal, 2), Array(BOM2(intIndexGlobal, 0), BOM2(intIndexGlobal, 1), BOM2(intIndexGlobal, 2), BOM2(intIndexGlobal, 3))
                       '---------加入全局字典以便Extract_PartList过程调用
                       If Not pDic.exists(oProducts1.Item(i).PartNumber) Then
                                Dim arrpDic(), h
'                                MsgBox UBound(BOM2, 2)
                                For h = 4 To UBound(BOM2, 2)
                                    ReDim Preserve arrpDic(h - 4)
                                    arrpDic(h - 4) = PropertyValue(oProducts1.Item(i), BOM2(0, h))
                                    BOM2(intIndexGlobal, h) = arrpDic(h - 4)   '加入当前BOM2数组
                                Next
                                pDic.Add oProducts1.Item(i).PartNumber, arrpDic
                       Else        'pDic已经存在这个零件的属性则直接读取
                                Dim T As Integer
                                If UBound(BOM2, 2) > 3 Then   '如果不大于3则字典键值为空数组，下标溢出
                                For T = 0 To UBound(pDic.Item(oProducts1.Item(i).PartNumber))
                                    BOM2(intIndexGlobal, T + 4) = pDic.Item(oProducts1.Item(i).PartNumber)(T)
                                Next
                                End If
                        End If
                    Else                                                                                       '----------------1
                       ProExist = pDicX.Item(oProducts1.Item(i).PartNumber)
                       ProExist(3) = ProExist(3) + 1
                       pDicX.Item(oProducts1.Item(i).PartNumber) = ProExist
                       
                       '反查回去修改哪一行
                            T = 0
                            Do Until (BOM2(intIndexGlobal - T, 2) = ProExist(2)) And (BOM2(intIndexGlobal - T, 1) = ProExist(1))
                            T = T + 1
                            Loop
                            BOM2(intIndexGlobal - T, 3) = ProExist(3)
                    End If                                                                                           '----------------1
                     ProExist = pDicX.Item(oProducts1.Item(i).PartNumber)
                        If (is_Leaf(oProducts1.Item(i)) = False) And ProExist(3) < 2 Then
                            Call Extract_BOM(oProducts1.Item(i))
                            intLevelGlobal = intLevelGlobal - 1              'when finished, intLevelGlobal minus one
                         End If
            Else  '零件未加载
                'MsgBox PN
                If Not pDicX.exists(PN) Then
                       intIndexGlobal = intIndexGlobal + 1
                       Dim ProExist1
                       BOM2(intIndexGlobal, 0) = intIndexGlobal 'Index
                       BOM2(intIndexGlobal, 1) = String(intLevelGlobal, "- ") & intLevelGlobal                      'Level
                       BOM2(intIndexGlobal, 2) = PN               'Same as dictory key
                       BOM2(intIndexGlobal, 3) = 1                                                                          'Qty
                       pDicX.Add BOM2(intIndexGlobal, 2), Array(BOM2(intIndexGlobal, 0), BOM2(intIndexGlobal, 1), BOM2(intIndexGlobal, 2), BOM2(intIndexGlobal, 3))
                    Else
                       ProExist1 = pDicX.Item(PN)
                       ProExist1(3) = ProExist1(3) + 1
                       pDicX.Item(PN) = ProExist1
                        '反查回去修改哪一行
                            T = 0
                            Do Until BOM2(intIndexGlobal - T, 2) = ProExist1(2) And (BOM2(intIndexGlobal - T, 1) = ProExist1(1))
                            T = T + 1
                            Loop
                            BOM2(intIndexGlobal - T, 3) = ProExist1(3)
                    End If
            End If '
    oProducts1.Item(i).ApplyWorkMode DEFAULT_MODE 'VISUALIZATION_MODE
     Next
End Sub
'*******************************提取零件清单********************************************
Sub Extract_PartList(ByVal oProduct1)
    Dim i As Integer
    Dim oProducts1
    Set oProducts1 = oProduct1.Products
     On Error Resume Next
    If is_Leaf(oProduct1) Then
         Exit Sub
    End If
    For i = 1 To oProducts1.Count
    oProducts1.Item(i).ApplyWorkMode DESIGN_MODE
                    Err.Clear
                    Dim Unload1, PN
                    PN = oProducts1.Item(i).PartNumber
                    If Err.Number <> 0 Then
                    Unload1 = True
                    PN = "未加载无法取得数据"
                    Else
                    Unload1 = False
                    End If
                    Err.Clear

                     If Unload1 = False Then
                     cmdRun1.Caption = "操作" & oProducts1.Item(i).PartNumber & "..."
                                If i Mod 2 = 0 Then
                                cmdRun1.BackColor = &H8000000D '&H8000000F
                                Else
                                cmdRun1.BackColor = &HFFFF&
                                End If
                         If is_Leaf(oProducts1.Item(i)) And (Not ProductIsComponent(oProducts1.Item(i))) Then                           '----------1
                                If Not oDicPL.exists(oProducts1.Item(i).PartNumber) Then
                                   intIndexPL = intIndexPL + 1
                                   Dim ProExist
                                   PL2(intIndexPL, 0) = intIndexPL                    'Index
                                   PL2(intIndexPL, 1) = Round(oProducts1.Item(i).Analyze.Mass * 1000, 2)                    'reserve for mass(xN)
                                   PL2(intIndexPL, 2) = oProducts1.Item(i).PartNumber               'Same as dictory key
                                   PL2(intIndexPL, 3) = 1
                                   oDicPL.Add PL2(intIndexPL, 2), Array(PL2(intIndexPL, 0), PL2(intIndexPL, 1), PL2(intIndexPL, 2), PL2(intIndexPL, 3))
                                    '---------调用Exact_BOM产生的全局字典pDic
                                                 Dim h1
                                                 If UBound(PL2, 2) > 3 Then
                                                 For h1 = 4 To UBound(PL2, 2)
                                                     PL2(intIndexPL, h1) = pDic.Item(oProducts1.Item(i).PartNumber)(h1 - 4)
                                                 Next
                                                 End If
                                Else
                                   ProExist = oDicPL.Item(oProducts1.Item(i).PartNumber)
                                   ProExist(3) = ProExist(3) + 1
                                   ProExist(1) = Round(oProducts1.Item(i).Analyze.Mass * 1000, 2) * ProExist(3)
                                   oDicPL.Item(oProducts1.Item(i).PartNumber) = ProExist
                                   
                                     '反查回去修改哪一行
                                        Dim T As Integer
                                             T = 0
                                             Do Until PL2(intIndexPL - T, 2) = ProExist(2)
                                             T = T + 1
                                             Loop
                                             PL2(intIndexPL - T, 3) = ProExist(3)
                                             PL2(intIndexPL - T, 1) = ProExist(1)
                                 End If
                        Else                                                                                                                                             '----------1
                                    Call Extract_PartList(oProducts1.Item(i))
                        End If                                                                                                                                           '----------1
                    Else  '零件未加载
                            If Not oDicPL.exists(PN) Then
                               intIndexPL = intIndexPL + 1
                               Dim ProExist1
                               PL2(intIndexPL, 0) = intIndexPL 'Index
                               PL2(intIndexPL, 1) = ""                   'mass(xN)
                               PL2(intIndexPL, 2) = PN              'Same as dictory key
                               PL2(intIndexPL, 3) = 1
                               oDicPL.Add PL2(intIndexPL, 2), Array(PL2(intIndexPL, 0), PL2(intIndexPL, 1), PL2(intIndexPL, 2), PL2(intIndexPL, 3))
                            Else
                               ProExist1 = oDicPL.Item(PN)
                               ProExist1(3) = ProExist1(3) + 1
                               oDicPL.Item(PN) = ProExist1
                                         '反查回去修改哪一行
                                             T = 0
                                             Do Until PL2(intIndexPL - T, 2) = ProExist1(2)
                                             T = T + 1
                                             Loop
                                             PL2(intIndexPL - T, 3) = ProExist1(3)
                                             PL2(intIndexPL - T, 1) = ProExist1(1)
                             End If
            End If
         oProducts1.Item(i).ApplyWorkMode DEFAULT_MODE
     Next
End Sub




'************************************输出当前Product用户属性名称，一维数列****************************************
Private Sub GetUserProperties(ByVal oProduct1)
On Error Resume Next
Dim oUserPara1 As Object 'Parameters
Set oUserPara1 = oProduct1.ReferenceProduct.UserRefProperties
If Err.Number <> 0 Then
Exit Sub
End If
'MsgBox UBound(UserProperties)
       If UBound(UserProperties) = -1 Then      '初始为空                    ------1#
            Dim n
            For n = 0 To oUserPara1.Count - 1
            ReDim Preserve UserProperties(n)
            UserProperties(n) = GetUserPropStr(oUserPara1.Item(n + 1).Name)
            Next
       Else                                                                                      '------1#
            Dim i As Integer
            For i = 0 To oUserPara1.Count - 1       '将此零件的用户属性与已经取得的用户属性名逐个比对
                Dim ExistFlag As Boolean
                ExistFlag = False
                Dim j As Integer
                             For j = 0 To UBound(UserProperties)
                                     If GetUserPropStr(oUserPara1.Item(i + 1).Name) = UserProperties(j) Then '当发现这个用户属性已经在用户数组名称中存在时
                                          ExistFlag = True
                                     End If
                             Next
                If ExistFlag = False Then
                       '否则用户属性名称数组增加这个名称， 数组长度加1
                            Dim xx As Integer
                            xx = UBound(UserProperties)
                            ReDim Preserve UserProperties(xx + 1)
                            UserProperties(xx + 1) = GetUserPropStr(oUserPara1.Item(i + 1).Name)
                            ' This VBA Macro Developed by Charles.Tang
                            ' WeChat Chtang80,CopyRight reserved
                End If
   
            Next
       End If                                                                                            '------1#

End Sub
'************************************输出当前Part或Product用户属性名称，一维数列****************************************
Private Sub GetUserProperties2(ByVal oProduct1)
GetUserProperties oProduct1
Dim ii
If oProduct1.Products.Count <> 0 Then
    For ii = 1 To oProduct1.Products.Count
       GetUserProperties2 oProduct1.Products.Item(ii)
    Next
End If

End Sub

''***************************隐藏基准面******************************
'Sub HideDatums(ByVal oProducti As Product)
'If ProductIsComponent(oProducti) Then
'Exit Sub
'End If
'On Error Resume Next
'CATIA.StatusBar = "正在截图，请稍侯..."
'If oProducti.Products.Count = 0 Then
'HidePartDatums oProducti.ReferenceProduct.Parent.Part
'Else
'Dim j
'    For j = 1 To oProducti.Products.Count
'    Call HideDatums(oProducti.Products.Item(j))
'    Next
'End If
'End Sub
'Sub HidePartDatums(oParti)
'Dim pSel As Selection
'Set pSel = oParti.Parent.Selection
'pSel.Clear
'Dim colAxSys, OrigEl
'Set colAxSys = oParti.AxisSystems
'Set OrigEl = oParti.OriginElements
'Dim i As Integer
'For i = 1 To colAxSys.Count
'pSel.Add colAxSys.Item(i)
'Next
'pSel.Add OrigEl.PlaneXY
'pSel.Add OrigEl.PlaneZX
'pSel.Add OrigEl.PlaneYZ
'' This VBA Macro Developed by Charles.Tang
'' WeChat Chtang80,CopyRight reserved
'Dim hybridBodies1 As Object
'Set hybridBodies1 = oParti.HybridBodies
'For i = 1 To hybridBodies1.Count
'pSel.Add hybridBodies1.Item(i)
'Next
'Dim bd As Object
'Dim sks As Object
'For i = 1 To oParti.Bodies.Count
'     Dim j As Integer
'     Set sks = oParti.Bodies.Item(i).Sketches
'     For j = 1 To sks.Count
'        pSel.Add sks.Item(j)
'     Next
'Next
'pSel.VisProperties.SetShow 1
'pSel.Clear
'End Sub
'***************************零件第一个用户轴系的方向******************************
Sub FVDirection(part1 As Object)
FrontViewDirection(0) = 0
FrontViewDirection(1) = 1
FrontViewDirection(2) = 0
FrontViewDirection(3) = 0
FrontViewDirection(4) = 0
FrontViewDirection(5) = 1

On Error Resume Next
    Err.Clear
            Dim axisSystems1 As Object
            Set axisSystems1 = part1.AxisSystems
            Dim axisSystem1 As Object
            Set axisSystem1 = axisSystems1.Item(1)
            
             Dim YAxisCoord(2)
             axisSystem1.GetYAxis YAxisCoord
             Dim ZAxisCoord(2)
             axisSystem1.GetZAxis ZAxisCoord
      If Err.Number = 0 Then
            FrontViewDirection(0) = YAxisCoord(0)
            FrontViewDirection(1) = YAxisCoord(1)
            FrontViewDirection(2) = YAxisCoord(2)
            FrontViewDirection(3) = ZAxisCoord(0)
            FrontViewDirection(4) = ZAxisCoord(1)
            FrontViewDirection(5) = ZAxisCoord(2)
       End If
On Error GoTo 0
End Sub
'***************************为产品及下属零件截图******************************
Sub CaptureChildrenImage(ByVal oProduct9, ByVal OutputFolder9)

CATIA.RefreshDisplay = False
On Error Resume Next
Dim HH, WW As Double  '当前窗口高度和宽度
HH = CATIA.ActiveWindow.Height
WW = CATIA.ActiveWindow.Width
CATIA.ActiveWindow.Height = 400  'v040 add
CATIA.ActiveWindow.Width = 600  'v040 add
Dim win1 As Object
Set win1 = CATIA.ActiveWindow
CATIA.StatusBar = "正在截图，请稍侯..."
Dim myViewer1
Dim color(2)
Set myViewer1 = CATIA.ActiveWindow.ActiveViewer

myViewer1.GetBackgroundColor color
myViewer1.PutBackgroundColor Array(1, 1, 1)
CATIA.ActiveWindow.Layout = 1

'CaptureWindowImage OutputFolder9 & "\" & oProduct9.PartNumber & ".jpg"
cmdRun1.Caption = "截图" & oProduct9.PartNumber & "..."
CaptureWindowImage OutputFolder9 & "\" & RebuildFileName(oProduct9.PartNumber) & ".jpg"

Dim selection9 As Selection
Set selection9 = CATIA.ActiveDocument.Selection
selection9.Clear
Dim visPropertySet9 As VisPropertyset
Set visPropertySet9 = selection9.VisProperties
Dim oProducts9
Set oProducts9 = oProduct9.Products

Dim i
For i = 1 To oProducts9.Count
selection9.Add oProducts9.Item(i)
Next
visPropertySet9.SetShow 1   '0-show,1-NoShow
selection9.Clear

For i = 1 To oProducts9.Count
selection9.Add oProducts9.Item(i)
visPropertySet9.SetShow 0
selection9.Clear
cmdRun1.Caption = "截图" & oProducts9.Item(i).PartNumber & "..."
Call CaptureChildrenImage(oProducts9.Item(i), OutputFolder9)
selection9.Add oProducts9.Item(i)
visPropertySet9.SetShow 1
selection9.Clear
Next


CATIA.ActiveWindow.Layout = 2 ' catWindowSpecsAndGeom
myViewer1.PutBackgroundColor color

'show all components

For i = 1 To oProducts9.Count
selection9.Add oProducts9.Item(i)
Next

visPropertySet9.SetShow 0

selection9.Clear
On Error GoTo 0

Set myViewer1 = Nothing
Set selection9 = Nothing
Set visPropertySet9 = Nothing
Set oProducts9 = Nothing
win1.Height = HH
win1.Width = WW
win1.Activate
Set win1 = Nothing
CATIA.StatusBar = "正在导出BOM，请稍侯..."
'CATIA.RefreshDisplay = True
End Sub
'***************************当前窗口截图******************************
Sub CaptureWindowImage(ImageFullName9 As String) ' ImageFullName9 end with .jpg
On Error Resume Next
CATIA.RefreshDisplay = False
Dim specsAndGeomWindow1 As SpecsAndGeomWindow
Set specsAndGeomWindow1 = CATIA.ActiveWindow

Dim viewer3D1 As Viewer3D
Set viewer3D1 = specsAndGeomWindow1.ActiveViewer
    viewer3D1.RenderingMode = 1 ' "catRenderShadingWithEdges"
    viewer3D1.Reframe
    viewer3D1.CaptureToFile 5, ImageFullName9 '"catCaptureFormatJPEG", ImageName
Set viewer3D1 = Nothing
Set specsAndGeomWindow1 = Nothing
'CATIA.RefreshDisplay = True
End Sub

Private Sub chkImg_Change()
If chkImg.Value = True Then
chkImgCol.Enabled = True
Else
chkImgCol.Value = False
chkImgCol.Enabled = False
End If
End Sub

Private Sub chkImgCol_Change()
If chkImgCol.Value = True Then
iImgCol = 1
Else
iImgCol = 0
End If
End Sub
'Sub hideConstraints(oProd As Object)
'Dim pSel As Selection
'Set pSel = CATIA.ActiveDocument.Selection
'pSel.Clear
'
'If oProd.Products.Count = 0 Then
'Exit Sub
'End If
'
'Dim oConstraints As Object
'Dim i2 As Integer
'Set oConstraints = oProd.Connections("CATIAConstraints")
'For i2 = 1 To oConstraints.Count
'pSel.Add oConstraints.Item(i2)
''Debug.Print oConstraints.Item(i2).Name & " hided"
'Next
'pSel.VisProperties.SetShow 1
'
'Dim oProdi As Object
'
'For Each oProdi In oProd.Products
''Debug.Print "处理产品" & oProdi.Name
'hideConstraints oProdi
'Next
'
'End Sub
Sub HideDatumsConstraints(prodoc As Document, intshow As Integer)

On Error Resume Next
Dim selection1 As Selection
Set selection1 = prodoc.Selection
selection1.Clear

selection1.Search "(((CATStFreeStyleSearch.Plane + CATPrtSearch.Plane) + CATGmoSearch.Plane) + CATSpdSearch.Plane),all"

selection1.VisProperties.SetShow intshow

selection1.Clear

selection1.Search "(((CATStFreeStyleSearch.AxisSystem + CATPrtSearch.AxisSystem) + CATGmoSearch.AxisSystem) + CATSpdSearch.AxisSystem),all"


selection1.VisProperties.SetShow intshow

selection1.Clear

selection1.Search "((((((CATStFreeStyleSearch.Point + CAT2DLSearch.2DPoint) + CATSketchSearch.2DPoint) + CATDrwSearch.2DPoint) + CATPrtSearch.Point) + CATGmoSearch.Point) + CATSpdSearch.Point),all"
selection1.VisProperties.SetShow intshow

selection1.Clear

selection1.Search "((((((CATStFreeStyleSearch.Curve + CAT2DLSearch.2DCurve) + CATSketchSearch.2DCurve) + CATDrwSearch.2DCurve) + CATPrtSearch.Curve) + CATGmoSearch.Curve) + CATSpdSearch.Curve),all"
selection1.VisProperties.SetShow intshow

selection1.Clear

selection1.Search "(((CATStFreeStyleSearch.Surface + CATPrtSearch.Surface) + CATGmoSearch.Surface) + CATSpdSearch.Surface),all"
selection1.VisProperties.SetShow intshow

selection1.Clear

selection1.Search "(((((((CATProductSearch.MfConstraint + CATStFreeStyleSearch.MfConstraint) + CATAsmSearch.MfConstraint) + CAT2DLSearch.MfConstraint) + CATSketchSearch.MfConstraint) + CATDrwSearch.MfConstraint) + CATPrtSearch.MfConstraint) + CATSpdSearch.MfConstraint),all"
selection1.VisProperties.SetShow intshow
selection1.Clear
Err.Clear
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
    read_OK = GetPrivateProfileString("导出BOM2", "为料号添加截图注释", "False", read2, 256, sConfpath & "\Conf1.ini")
    chkImg.Value = read2
    read2 = String(255, 0)
    read_OK = GetPrivateProfileString("导出BOM2", "添加一列零件截图", "False", read2, 256, sConfpath & "\Conf1.ini")
    chkImgCol.Value = read2

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
write1 = WritePrivateProfileString("导出BOM2", "为料号添加截图注释", CStr(chkImg.Value), sConfpath & "\Conf1.ini")
write1 = WritePrivateProfileString("导出BOM2", "添加一列零件截图", CStr(chkImgCol.Value), sConfpath & "\Conf1.ini")
End Sub

Private Sub UserForm_Terminate()
SaveConf
End Sub
Private Function RebuildFileName(s As String)
Dim re As Object

Set re = CreateObject("Vbscript.RegExp")
re.Pattern = "[^A-Za-z0-9_ (\u4e00-\u9fa5)-]"
re.Global = True
RebuildFileName = re.Replace(s, "_")
Set re = Nothing
End Function
